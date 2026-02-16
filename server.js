import express from 'express';
import { chromium } from 'playwright';
import Anthropic from '@anthropic-ai/sdk';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import multer from 'multer';
import { parse } from 'csv-parse/sync';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
app.use(express.json());
app.use(express.static(join(__dirname, 'public')));

// In-memory store for completed CSV results (keyed by sessionId)
const sessionResults = new Map();

// Multer: store CSV in memory (no disk write needed)
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });

// ─── Playwright Action Executor ───────────────────────────────────────────────

async function executeAction(page, action) {
  const { action: actionType, coordinate, text, key, scroll_direction, scroll_amount } = action;

  switch (actionType) {
    case 'screenshot': break; // handled separately
    case 'left_click': {
      const [x, y] = coordinate;
      await page.mouse.click(x, y);
      break;
    }
    case 'right_click': {
      const [x, y] = coordinate;
      await page.mouse.click(x, y, { button: 'right' });
      break;
    }
    case 'double_click': {
      const [x, y] = coordinate;
      await page.mouse.dblclick(x, y);
      break;
    }
    case 'triple_click': {
      const [x, y] = coordinate;
      await page.mouse.click(x, y, { clickCount: 3 });
      break;
    }
    case 'middle_click': {
      const [x, y] = coordinate;
      await page.mouse.click(x, y, { button: 'middle' });
      break;
    }
    case 'mouse_move': {
      const [x, y] = coordinate;
      await page.mouse.move(x, y);
      break;
    }
    case 'left_click_drag': {
      const [sx, sy] = action.start_coordinate || coordinate;
      const [ex, ey] = coordinate;
      await page.mouse.move(sx, sy);
      await page.mouse.down();
      await page.mouse.move(ex, ey);
      await page.mouse.up();
      break;
    }
    case 'type': {
      await page.keyboard.type(text);
      break;
    }
    case 'key': {
      const keyMap = {
        'Return': 'Enter', 'ctrl+a': 'Control+a', 'ctrl+c': 'Control+c',
        'ctrl+v': 'Control+v', 'ctrl+z': 'Control+z', 'super': 'Meta',
      };
      await page.keyboard.press(keyMap[key] || key);
      break;
    }
    case 'scroll': {
      const [x, y] = coordinate;
      const amount = scroll_amount || 3;
      const deltaX = scroll_direction === 'left' ? -amount * 100 : scroll_direction === 'right' ? amount * 100 : 0;
      const deltaY = scroll_direction === 'up' ? -amount * 100 : scroll_direction === 'down' ? amount * 100 : 0;
      await page.mouse.wheel(deltaX, deltaY);
      break;
    }
    case 'wait': {
      await page.waitForTimeout(action.duration ? action.duration * 1000 : 1000);
      break;
    }
    default:
      console.log(`Unknown action: ${actionType}`);
  }
}

// ─── Take Screenshot ──────────────────────────────────────────────────────────

async function takeScreenshot(page) {
  const buffer = await page.screenshot({ type: 'png' });
  return buffer.toString('base64');
}

// ─── SSE Helper ──────────────────────────────────────────────────────────────

function sendEvent(res, type, data) {
  res.write(`data: ${JSON.stringify({ type, ...data })}\n\n`);
}

// ─── Extract Reference Number from page text ─────────────────────────────────

async function extractReferenceNumber(page, claudeTexts) {
  // Strategy 1: scan page DOM for ref number patterns
  try {
    const pageText = await page.evaluate(() => document.body.innerText);
    // Common patterns: "Reference number: ABC123", "Ref: ABC123", "Reference: ABC123"
    const patterns = [
      /reference\s*(?:number|#|num|no\.?)?\s*[:\-]?\s*([A-Z0-9]{6,})/i,
      /ref(?:erence)?\s*(?:number|#|num|no\.?)?\s*[:\-]\s*([A-Z0-9]{6,})/i,
      /confirmation\s*(?:number|#|code)?\s*[:\-]?\s*([A-Z0-9]{6,})/i,
      /transaction\s*(?:id|#)?\s*[:\-]?\s*([A-Z0-9]{6,})/i,
      /ticket\s*(?:number|#)?\s*[:\-]?\s*([A-Z0-9]{6,})/i,
      /submission\s*(?:id|number|#)?\s*[:\-]?\s*([A-Z0-9]{6,})/i,
      /\b([A-Z0-9]{8,})\b/,  // fallback: any long alphanumeric token
    ];
    for (const pat of patterns) {
      const m = pageText.match(pat);
      if (m) return m[1].trim();
    }
  } catch (_) {}

  // Strategy 2: scan Claude's text responses for the reference number
  const allText = claudeTexts.join(' ');
  const claudePatterns = [
    /reference\s*(?:number|#|num|no\.?)?\s*[:\-]?\s*([A-Z0-9]{6,})/i,
    /ref(?:erence)?\s*[:\-]\s*([A-Z0-9]{6,})/i,
    /confirmation\s*(?:number|#|code)?\s*[:\-]?\s*([A-Z0-9]{6,})/i,
    /\b([A-Z0-9]{8,})\b/,
  ];
  for (const pat of claudePatterns) {
    const m = allText.match(pat);
    if (m) return m[1].trim();
  }

  return 'N/A';
}

// ─── Run one form-fill agent loop for a single row ───────────────────────────

async function runFormFillForRow(client, page, url, fieldsObj, sendEvent, res, rowIndex, totalRows) {
  // Navigate/reload the form URL fresh for each row
  sendEvent(res, 'row_status', { rowIndex, message: `Navigating to form for row ${rowIndex + 1}...` });
  await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
  await page.waitForTimeout(1500);

  const fieldsList = Object.entries(fieldsObj)
    .map(([k, v]) => `  - "${k}": "${v}"`)
    .join('\n');

  const task = `You are automating a web browser to fill out a form. The browser is already open at the correct URL.

Your task: Fill out the form on this page with the following data, then submit it:
${fieldsList}

Instructions:
1. Take a screenshot first to see the current state of the page.
2. Identify each form field and fill it with the provided data.
3. After filling all fields, click the Submit button.
4. After submission, take a screenshot of the confirmation/thank-you page.
5. Look for a reference number, confirmation number, or ticket number on the confirmation page and state it clearly in your final message as "Reference number: XXXXX".
6. Be precise with coordinates. Click directly on input fields before typing.
7. After each major action, take a screenshot to verify success.`;

  const messages = [{ role: 'user', content: task }];
  const tools = [{
    type: 'computer_20250124',
    name: 'computer',
    display_width_px: 1280,
    display_height_px: 800,
  }];

  let iteration = 0;
  const MAX_ITERATIONS = 25;
  let taskComplete = false;
  const claudeTexts = [];

  while (iteration < MAX_ITERATIONS && !taskComplete) {
    iteration++;
    sendEvent(res, 'iteration', { iteration, max: MAX_ITERATIONS, rowIndex, totalRows });

    const response = await client.beta.messages.create({
      model: 'claude-sonnet-4-5-20250929',
      max_tokens: 4096,
      tools,
      messages,
      betas: ['computer-use-2025-01-24'],
    });

    messages.push({ role: 'assistant', content: response.content });

    const toolResults = [];
    let hasToolUse = false;

    for (const block of response.content) {
      if (block.type === 'text') {
        if (block.text) claudeTexts.push(block.text);
        sendEvent(res, 'claude_text', { text: block.text, rowIndex });
      } else if (block.type === 'tool_use' && block.name === 'computer') {
        hasToolUse = true;
        const action = block.input;
        sendEvent(res, 'action', { action: action.action, details: action, rowIndex });

        let toolResultContent;

        if (action.action === 'screenshot') {
          const screenshotB64 = await takeScreenshot(page);
          sendEvent(res, 'screenshot', { image: screenshotB64, rowIndex });
          toolResultContent = [{
            type: 'image',
            source: { type: 'base64', media_type: 'image/png', data: screenshotB64 },
          }];
        } else {
          try {
            await executeAction(page, action);
            await page.waitForTimeout(500);
            const screenshotB64 = await takeScreenshot(page);
            sendEvent(res, 'screenshot', { image: screenshotB64, rowIndex });
            toolResultContent = [{
              type: 'image',
              source: { type: 'base64', media_type: 'image/png', data: screenshotB64 },
            }];
          } catch (err) {
            toolResultContent = `Error executing action ${action.action}: ${err.message}`;
            sendEvent(res, 'error', { message: toolResultContent, rowIndex });
          }
        }

        toolResults.push({
          type: 'tool_result',
          tool_use_id: block.id,
          content: toolResultContent,
        });
      }
    }

    if (response.stop_reason === 'end_turn' && !hasToolUse) {
      taskComplete = true;
      break;
    }
    if (!hasToolUse) {
      taskComplete = true;
      break;
    }

    messages.push({ role: 'user', content: toolResults });
  }

  // Take final screenshot and extract reference number
  const finalScreenshot = await takeScreenshot(page);
  sendEvent(res, 'screenshot', { image: finalScreenshot, rowIndex, isFinal: true });

  const refNumber = await extractReferenceNumber(page, claudeTexts);
  return { success: taskComplete, iterations: iteration, refNumber };
}

// ─── CSV Upload + Batch Automate Endpoint (SSE) ───────────────────────────────

app.post('/api/automate-csv', upload.single('csvFile'), async (req, res) => {
  const { url, apiKey, sessionId } = req.body;

  if (!url || !apiKey || !sessionId) {
    return res.status(400).json({ error: 'Missing required fields: url, apiKey, sessionId' });
  }
  if (!req.file) {
    return res.status(400).json({ error: 'No CSV file uploaded' });
  }

  // Parse CSV
  let rows;
  try {
    rows = parse(req.file.buffer.toString('utf8'), {
      columns: true,       // use first row as header
      skip_empty_lines: true,
      trim: true,
    });
  } catch (err) {
    return res.status(400).json({ error: `CSV parse error: ${err.message}` });
  }

  if (rows.length === 0) {
    return res.status(400).json({ error: 'CSV file is empty or has no data rows' });
  }

  // Set up SSE
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.flushHeaders();

  const client = new Anthropic({ apiKey });
  let browser = null;

  // Initialise results store for this session
  const headers = Object.keys(rows[0]);
  const results = rows.map(r => ({ ...r, ReferenceNumber: '' }));
  sessionResults.set(sessionId, { headers, rows: results, originalName: req.file.originalname });

  sendEvent(res, 'batch_start', { total: rows.length, headers });

  try {
    browser = await chromium.launch({ headless: false, args: ['--no-sandbox'] });
    const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
    const page = await context.newPage();

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      sendEvent(res, 'row_start', { rowIndex: i, total: rows.length, data: row });

      try {
        const { success, iterations, refNumber } = await runFormFillForRow(
          client, page, url, row, sendEvent, res, i, rows.length
        );

        results[i].ReferenceNumber = refNumber;
        sessionResults.get(sessionId).rows = results;  // keep store updated

        sendEvent(res, 'row_done', {
          rowIndex: i,
          total: rows.length,
          success,
          iterations,
          refNumber,
          data: row,
        });
      } catch (err) {
        results[i].ReferenceNumber = 'ERROR';
        sessionResults.get(sessionId).rows = results;
        sendEvent(res, 'row_error', { rowIndex: i, message: err.message });
      }

      // Brief pause between rows
      if (i < rows.length - 1) {
        await page.waitForTimeout(1500);
      }
    }

    sendEvent(res, 'batch_done', { total: rows.length, sessionId });

  } catch (err) {
    console.error('Batch automation error:', err);
    sendEvent(res, 'error', { message: err.message });
    sendEvent(res, 'batch_done', { total: rows.length, sessionId, error: err.message });
  } finally {
    if (browser) setTimeout(() => browser.close(), 4000);
    res.end();
  }
});

// ─── Download Results CSV ─────────────────────────────────────────────────────

app.get('/api/download-results/:sessionId', (req, res) => {
  const { sessionId } = req.params;
  const data = sessionResults.get(sessionId);
  if (!data) return res.status(404).json({ error: 'Session not found or expired' });

  const { headers, rows } = data;
  const allHeaders = [...headers, 'ReferenceNumber'];

  // Build CSV string manually (no external dep needed for output)
  const escape = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const lines = [
    allHeaders.map(escape).join(','),
    ...rows.map(r => allHeaders.map(h => escape(r[h] ?? '')).join(',')),
  ];
  const csv = lines.join('\r\n');

  const filename = 'reference-numbers.csv';
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(csv);
});

// ─── Single-row Automate Endpoint (SSE) — unchanged ──────────────────────────

app.get('/api/automate', async (req, res) => {
  const { url, fields, apiKey } = req.query;

  if (!url || !fields || !apiKey) {
    return res.status(400).json({ error: 'Missing required parameters: url, fields, apiKey' });
  }

  let parsedFields;
  try {
    parsedFields = JSON.parse(fields);
  } catch {
    return res.status(400).json({ error: 'Invalid fields JSON' });
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.flushHeaders();

  const client = new Anthropic({ apiKey });
  let browser = null;

  try {
    sendEvent(res, 'status', { message: 'Launching browser...' });
    browser = await chromium.launch({ headless: false, args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
    const page = await context.newPage();

    sendEvent(res, 'status', { message: `Navigating to ${url}` });
    await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
    await page.waitForTimeout(1500);

    const fieldsList = Object.entries(parsedFields).map(([k, v]) => `  - "${k}": "${v}"`).join('\n');
    const task = `You are automating a web browser to fill out a form. The browser is already open at the correct URL.

Your task: Fill out the form on this page with the following data, then submit it:
${fieldsList}

Instructions:
1. Take a screenshot first to see the current state of the page.
2. Identify each form field and fill it with the provided data.
3. After filling all fields, click the Submit button.
4. Confirm submission was successful.
5. Be precise with coordinates. Click directly on input fields before typing.
6. After each major action, take a screenshot to verify success.`;

    const messages = [{ role: 'user', content: task }];
    const tools = [{ type: 'computer_20250124', name: 'computer', display_width_px: 1280, display_height_px: 800 }];

    let iteration = 0;
    const MAX_ITERATIONS = 25;
    let taskComplete = false;

    sendEvent(res, 'status', { message: 'Starting Claude computer-use agent loop...' });

    while (iteration < MAX_ITERATIONS && !taskComplete) {
      iteration++;
      sendEvent(res, 'iteration', { iteration, max: MAX_ITERATIONS });
      sendEvent(res, 'status', { message: `Iteration ${iteration}: Calling Claude...` });

      const response = await client.beta.messages.create({
        model: 'claude-sonnet-4-5-20250929',
        max_tokens: 4096,
        tools,
        messages,
        betas: ['computer-use-2025-01-24'],
      });

      messages.push({ role: 'assistant', content: response.content });

      const toolResults = [];
      let hasToolUse = false;

      for (const block of response.content) {
        if (block.type === 'text') {
          sendEvent(res, 'claude_text', { text: block.text });
        } else if (block.type === 'tool_use' && block.name === 'computer') {
          hasToolUse = true;
          const action = block.input;
          sendEvent(res, 'action', { action: action.action, details: action });

          let toolResultContent;
          if (action.action === 'screenshot') {
            sendEvent(res, 'status', { message: 'Taking screenshot...' });
            const screenshotB64 = await takeScreenshot(page);
            sendEvent(res, 'screenshot', { image: screenshotB64 });
            toolResultContent = [{ type: 'image', source: { type: 'base64', media_type: 'image/png', data: screenshotB64 } }];
          } else {
            try {
              await executeAction(page, action);
              await page.waitForTimeout(500);
              const screenshotB64 = await takeScreenshot(page);
              sendEvent(res, 'screenshot', { image: screenshotB64 });
              toolResultContent = [{ type: 'image', source: { type: 'base64', media_type: 'image/png', data: screenshotB64 } }];
            } catch (err) {
              toolResultContent = `Error executing action ${action.action}: ${err.message}`;
              sendEvent(res, 'error', { message: toolResultContent });
            }
          }

          toolResults.push({ type: 'tool_result', tool_use_id: block.id, content: toolResultContent });
        }
      }

      if (response.stop_reason === 'end_turn' && !hasToolUse) { taskComplete = true; sendEvent(res, 'status', { message: 'Claude completed the task!' }); break; }
      if (!hasToolUse) { taskComplete = true; break; }
      messages.push({ role: 'user', content: toolResults });
    }

    if (!taskComplete && iteration >= MAX_ITERATIONS)
      sendEvent(res, 'status', { message: `Reached max iterations (${MAX_ITERATIONS}). Stopping.` });

    const finalScreenshot = await takeScreenshot(page);
    sendEvent(res, 'screenshot', { image: finalScreenshot, isFinal: true });
    sendEvent(res, 'done', { success: taskComplete, iterations: iteration });

  } catch (err) {
    console.error('Automation error:', err);
    sendEvent(res, 'error', { message: err.message });
    sendEvent(res, 'done', { success: false, error: err.message });
  } finally {
    if (browser) setTimeout(() => browser.close(), 5000);
    res.end();
  }
});

// ─── Start Server ─────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n🤖 Claude Form Automator running at http://localhost:${PORT}\n`);
});
