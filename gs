/*  GPT‑4o chat box for Google Sheets
 *  Sheet required:  "Chat"
 *  Named ranges required:  BYO_Key, Input, Reply, Run
 *  Use a drawing/button assigned to runChat() for the “Run” box.
 *  History is stored in DocumentProperties under key "HISTORY"
 *  Last 20 turns are sent for context each time.
 */

// ============ MAIN ============
function runChat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== 'Chat') {
    SpreadsheetApp.getUi().alert('Please run the chat from the "Chat" sheet.');
    return;
  }

  const keyCell    = ss.getRangeByName('BYO_Key');
  const promptCell = ss.getRangeByName('Input');
  const replyCell  = ss.getRangeByName('Reply');

  const apiKey = String(keyCell.getValue()).trim();
  if (!apiKey) {
    replyCell.setValue('🔑  Add your OpenAI key in BYO_Key and try again.');
    return;
  }

  const prompt = String(promptCell.getValue()).trim();
  if (!prompt) return;

  const props   = PropertiesService.getDocumentProperties();
  let history   = JSON.parse(props.getProperty('HISTORY') || '[]');
  history = history.slice(-20);
  history.push({ role: 'user', content: prompt });

  const payload = { model: 'gpt-4o-mini', messages: history, temperature: 0.7 };

  try {
    const res   = UrlFetchApp.fetch(
      'https://api.openai.com/v1/chat/completions',
      { method:'post', contentType:'application/json',
        headers:{ Authorization:'Bearer '+apiKey },
        payload: JSON.stringify(payload), muteHttpExceptions:true }
    );
    const body  = JSON.parse(res.getContentText());
    const reply = body.choices?.[0]?.message?.content?.trim() ||
                  `⚠️ No reply (status ${res.getResponseCode()})`;

    replyCell.setValue(reply);
    promptCell.setValue('');

    history.push({ role:'assistant', content: reply });
    props.setProperty('HISTORY', JSON.stringify(history));
  } catch (err) {
    replyCell.setValue('🚨  API error: ' + err.message);
  }
}

// ============ MENUS & UTILITIES ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Chat')
    .addItem('Send', 'runChat')
    .addSeparator()
    .addItem('Reset history', 'resetChat')
    .addToUi();
}

function resetChat() {
  PropertiesService.getDocumentProperties().deleteProperty('HISTORY');
  SpreadsheetApp.getActiveSpreadsheet()
                .getRangeByName('Reply')
                .setValue('🤖 History cleared.');
}

/*  onEdit lets you type anything in the Run cell and press ↵ */
function onEdit(e) {
  const ss = e.source;
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Chat') return;

  const runCell = ss.getRangeByName('Run');
  if (e.range.getA1Notation() === runCell.getA1Notation()) {
    runChat();
    runCell.setValue('');
  }
}
