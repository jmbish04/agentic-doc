/* === CONFIG & MENU === */

const DEFAULT_SETTINGS = {
  apiBase: "https://openai-assistant-proxy.hacolby.workers.dev/v1", // your proxy; switch to https://api.openai.com/v1 if desired
  apiKey: "",                 // store securely in Settings (User Properties)
  model: "gpt-4o-mini",       // any chat-completions-capable model
  systemPrompt:
`You are a helpful Google Docs editing agent.
You have tools to modify the active document: insert/replace text, add headings, insert images by URL, and create tables.
Prefer using tools for changes. If you need images, call insert_image_from_url.
When editing, keep user intent, formatting, and style in mind.`
};

function onOpen() {
  DocumentApp.getUi()
    .createMenu("AI Assistant")
    .addItem("Open Sidebar", "showSidebar")
    .addItem("Settings", "openSettings")
    .addToUi();
}

function showSidebar() {
  const t = HtmlService.createTemplateFromFile("Sidebar");
  t.settings = getSettings(); // pass current settings to client
  const html = t.evaluate().setTitle("AI Assistant");
  DocumentApp.getUi().showSidebar(html);
}

function openSettings() {
  const t = HtmlService.createTemplateFromFile("Sidebar");
  t.settings = getSettings();
  const html = t.evaluate().setTitle("AI Assistant");
  DocumentApp.getUi().showSidebar(html);
}

/* === SETTINGS STORAGE === */

function getSettings() {
  const userProps = PropertiesService.getUserProperties();
  const s = userProps.getProperty("OPENAI_SIDEBAR_SETTINGS");
  if (s) {
    try { return JSON.parse(s); } catch (_) {}
  }
  // First time: prefill defaults
  return DEFAULT_SETTINGS;
}

function saveSettings(partial) {
  const merged = Object.assign({}, getSettings(), partial || {});
  PropertiesService.getUserProperties().setProperty("OPENAI_SIDEBAR_SETTINGS", JSON.stringify(merged));
  return { ok: true };
}

/* === CHAT STATE (per-document) === */

function getDocState_() {
  const docProps = PropertiesService.getDocumentProperties();
  const s = docProps.getProperty("OPENAI_SIDEBAR_STATE");
  if (!s) return { messages: [] };
  try { return JSON.parse(s); } catch (_) { return { messages: [] }; }
}

function saveDocState_(state) {
  PropertiesService.getDocumentProperties()
    .setProperty("OPENAI_SIDEBAR_STATE", JSON.stringify(state || { messages: [] }));
}

/* === PUBLIC SERVER METHODS FOR CLIENT === */

function resetConversation() {
  saveDocState_({ messages: [] });
  return { ok: true };
}

function serverSendMessage(payload) {
  // payload: { text: string, allowEdits: boolean }
  const settings = getSettings();
  if (!settings.apiKey) throw new Error("Missing API key in Settings.");
  if (!settings.apiBase) throw new Error("Missing API base URL in Settings.");

  const state = getDocState_();

  // Build initial messages (system + history + user)
  const messages = [];
  if (settings.systemPrompt && settings.systemPrompt.trim()) {
    messages.push({ role: "system", content: settings.systemPrompt });
  }
  state.messages.forEach(m => messages.push(m)); // prior turns
  messages.push({ role: "user", content: payload.text });

  // Request loop: handle tool calls and final assistant text
  const toolDefs = getToolDefinitions_();  // list of tools we advertise
  const runResult = runChatWithTools_(settings, messages, toolDefs, payload.allowEdits);

  // Persist updated history
  saveDocState_({ messages: runResult.history });

  return {
    ok: true,
    finalText: runResult.finalText,
    actions: runResult.actionsApplied
  };
}

/* === OPENAI CHAT + TOOL CALL LOOP === */

function runChatWithTools_(settings, initialMessages, toolDefs, allowEdits) {
  const history = initialMessages.slice();
  const actionsApplied = [];

  // Safety: cap iterations to avoid infinite loops
  for (let step = 0; step < 6; step++) {
    const response = openAiChat_(settings, history, toolDefs);
    const choice = response && response.choices && response.choices[0];
    if (!choice) throw new Error("No choices returned from model.");

    const msg = choice.message || {};
    // If tool calls present, execute them (optionally) and return results as tool messages
    const toolCalls = (msg.tool_calls || []).filter(tc => tc && tc.type === "function");
    if (toolCalls.length) {
      history.push({ role: "assistant", content: msg.content || "", tool_calls: toolCalls });

      // Execute each function call
      const toolMessages = [];
      for (const tc of toolCalls) {
        const fnName = tc.function.name;
        const args = safeJsonParse_(tc.function.arguments);
        let result = { ok: true, note: "Tool disabled: allowEdits=false" };

        if (allowEdits) {
          result = executeToolCall_(fnName, args);
          actionsApplied.push({ name: fnName, args, result });
        } else {
          actionsApplied.push({ name: fnName, args, result });
        }

        toolMessages.push({
          role: "tool",
          tool_call_id: tc.id || undefined,
          name: fnName,
          content: JSON.stringify(result)
        });
      }

      // Send tool results back to the model for a follow-up turn
      history.push(...toolMessages);
      continue;
    }

    // No tool calls; we have final assistant content this step
    if (msg.content && msg.content.trim()) {
      history.push({ role: "assistant", content: msg.content });
      return { history, finalText: msg.content, actionsApplied };
    } else {
      // Assistant returned empty content and no tool calls; end
      return { history, finalText: "", actionsApplied };
    }
  }

  // Loop ended due to cap
  return { history, finalText: "(Stopped after tool-call steps limit.)", actionsApplied };
}

function openAiChat_(settings, messages, toolDefs) {
  const url = (settings.apiBase || "").replace(/\/+$/, "") + "/chat/completions";
  const payload = {
    model: settings.model || "gpt-4o-mini",
    messages: messages,
    temperature: 0.2,
    // Advertise our function tools to the model
    tools: toolDefs,
    tool_choice: "auto"
  };

  const headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + settings.apiKey
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    payload: JSON.stringify(payload),
    headers,
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("OpenAI error " + code + ": " + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

/* === TOOL DEFINITIONS (JSON Schema for function calling) === */

function getToolDefinitions_() {
  return [
    {
      type: "function",
      function: {
        name: "insert_text",
        description: "Insert plain text at start, cursor, or end of the document.",
        parameters: {
          type: "object",
          properties: {
            text: { type: "string" },
            location: { type: "string", enum: ["start", "cursor", "end"], default: "cursor" }
          },
          required: ["text"]
        }
      }
    },
    {
      type: "function",
      function: {
        name: "replace_text",
        description: "Find and replace text across the document body using a simple string or regex.",
        parameters: {
          type: "object",
          properties: {
            find: { type: "string", description: "Literal string or regex (surrounded by /.../)." },
            replace: { type: "string" },
            useRegex: { type: "boolean", default: false }
          },
          required: ["find", "replace"]
        }
      }
    },
    {
      type: "function",
      function: {
        name: "insert_heading",
        description: "Insert a new heading paragraph.",
        parameters: {
          type: "object",
          properties: {
            text: { type: "string" },
            level: { type: "integer", enum: [1,2,3,4,5,6], default: 2 },
            location: { type: "string", enum: ["start", "cursor", "end"], default: "cursor" }
          },
          required: ["text"]
        }
      }
    },
    {
      type: "function",
      function: {
        name: "insert_image_from_url",
        description: "Fetch an image by URL and insert it into the document.",
        parameters: {
          type: "object",
          properties: {
            url: { type: "string" },
            altText: { type: "string" },
            width: { type: "integer", minimum: 1 },
            height: { type: "integer", minimum: 1 },
            location: { type: "string", enum: ["start", "cursor", "end"], default: "end" }
          },
          required: ["url"]
        }
      }
    },
    {
      type: "function",
      function: {
        name: "insert_table",
        description: "Insert a table with optional cell contents.",
        parameters: {
          type: "object",
          properties: {
            rows: { type: "integer", minimum: 1 },
            cols: { type: "integer", minimum: 1 },
            data: {
              type: "array",
              items: { type: "array", items: { type: "string" } },
              description: "2D array [rows][cols]; fill missing with empty strings."
            },
            location: { type: "string", enum: ["start", "cursor", "end"], default: "end" }
          },
          required: ["rows", "cols"]
        }
      }
    }
  ];
}

/* === TOOL IMPLEMENTATIONS === */

function executeToolCall_(name, args) {
  try {
    switch (name) {
      case "insert_text": return toolInsertText_(args);
      case "replace_text": return toolReplaceText_(args);
      case "insert_heading": return toolInsertHeading_(args);
      case "insert_image_from_url": return toolInsertImageFromUrl_(args);
      case "insert_table": return toolInsertTable_(args);
      default: return { ok: false, error: "Unknown tool: " + name };
    }
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/* Helpers */

function getTargetInsertionPoint_(location) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  if (location === "start") return { doc, body, atStart: true, atEnd: false, cursor: null };
  if (location === "end")   return { doc, body, atStart: false, atEnd: true, cursor: null };
  const cursor = doc.getCursor();
  return { doc, body, atStart: false, atEnd: false, cursor };
}

function toolInsertText_(args) {
  const text = args.text || "";
  const loc = args.location || "cursor";
  const { body, cursor, atStart, atEnd } = getTargetInsertionPoint_(loc);

  let paragraph;
  if (cursor) {
    const el = cursor.insertText(text);
    if (!el) {
      paragraph = body.appendParagraph(text);
    }
  } else if (atStart) {
    paragraph = body.insertParagraph(0, text);
  } else if (atEnd) {
    paragraph = body.appendParagraph(text);
  } else {
    // No cursor available; default to end
    paragraph = body.appendParagraph(text);
  }

  DocumentApp.getActiveDocument().saveAndClose();
  return { ok: true, inserted: text.length };
}

function toolReplaceText_(args) {
  const find = args.find || "";
  const replace = args.replace || "";
  const useRegex = !!args.useRegex;

  const body = DocumentApp.getActiveDocument().getBody();
  let count = 0;

  if (useRegex || (find.startsWith("/") && find.endsWith("/"))) {
    // Regex path
    const pattern = useRegex ? find : find.slice(1, -1);
    const res = body.replaceText(pattern, replace);
    // replaceText returns Element; we can’t count replacements precisely, note success
    DocumentApp.getActiveDocument().saveAndClose();
    return { ok: true, note: "Regex replace applied." };
  } else {
    // Simple literal replace: iterate all text elements
    const found = body.findText(find);
    if (!found) {
      return { ok: true, replaced: 0, note: "No matches." };
    }
    let range = found;
    while (range) {
      const el = range.getElement();
      const start = range.getStartOffset();
      const end = range.getEndOffsetInclusive();
      const txt = el.asText();
      txt.deleteText(start, end);
      txt.insertText(start, replace);
      count++;
      range = body.findText(find, range);
    }
    DocumentApp.getActiveDocument().saveAndClose();
    return { ok: true, replaced: count };
  }
}

function toolInsertHeading_(args) {
  const text = args.text || "";
  const level = Math.max(1, Math.min(6, args.level || 2));
  const loc = args.location || "cursor";
  const { body, cursor, atStart, atEnd } = getTargetInsertionPoint_(loc);
  const headingMap = {
    1: DocumentApp.ParagraphHeading.HEADING1,
    2: DocumentApp.ParagraphHeading.HEADING2,
    3: DocumentApp.ParagraphHeading.HEADING3,
    4: DocumentApp.ParagraphHeading.HEADING4,
    5: DocumentApp.ParagraphHeading.HEADING5,
    6: DocumentApp.ParagraphHeading.HEADING6
  };

  let p;
  if (cursor) {
    p = cursor.insertText(text);
    if (!p) p = body.appendParagraph(text);
    p.getParent().asParagraph().setHeading(headingMap[level]);
  } else if (atStart) {
    p = body.insertParagraph(0, text).setHeading(headingMap[level]);
  } else {
    p = body.appendParagraph(text).setHeading(headingMap[level]);
  }

  DocumentApp.getActiveDocument().saveAndClose();
  return { ok: true, heading: level };
}

function toolInsertImageFromUrl_(args) {
  const url = args.url;
  if (!url) throw new Error("Missing image URL.");
  const alt = args.altText || "";
  const loc = args.location || "end";
  const w = args.width || null;
  const h = args.height || null;

  const { body, cursor, atStart, atEnd } = getTargetInsertionPoint_(loc);

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) throw new Error("Failed to fetch image, code " + code);
  const blob = resp.getBlob();

  let para;
  if (cursor) {
    para = cursor.insertInlineImage(blob);
    if (!para) para = body.appendImage(blob);
  } else if (atStart) {
    const p = body.insertParagraph(0, "");
    para = p.insertInlineImage(0, blob);
  } else if (atEnd) {
    para = body.appendImage(blob);
  }

  if (para && para.setAltDescription) {
    try { para.setAltDescription(alt); } catch (_){}
  }
  if (para && para.setWidth && w) para.setWidth(w);
  if (para && para.setHeight && h) para.setHeight(h);

  DocumentApp.getActiveDocument().saveAndClose();
  return { ok: true, note: "Image inserted." };
}

function toolInsertTable_(args) {
  const rows = Math.max(1, args.rows || 1);
  const cols = Math.max(1, args.cols || 1);
  const data = (args.data || []);
  const loc = args.location || "end";

  const { body, cursor, atStart } = getTargetInsertionPoint_(loc);

  // Build table
  const table = body.appendTable(); // default to end; adjust if start/cursor
  if (atStart) {
    body.removeChild(table);
    const t2 = body.insertTable(0, []);
    table.copyAttributesFrom(t2);
  }

  // Fill rows
  for (let r = 0; r < rows; r++) {
    const row = table.appendTableRow();
    for (let c = 0; c < cols; c++) {
      const cell = row.appendTableCell("");
      const v = (data[r] && data[r][c]) ? String(data[r][c]) : "";
      cell.setText(v);
    }
  }

  DocumentApp.getActiveDocument().saveAndClose();
  return { ok: true, rows, cols };
}

/* === UTIL === */

function safeJsonParse_(s) {
  try { return JSON.parse(s || "{}"); } catch (e) { return {}; }
}

/* === SERVER-SIDE SETTINGS BRIDGE === */

function getServerSettings() {
  return getSettings();
}
