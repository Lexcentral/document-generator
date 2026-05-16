// ── Web app endpoints ─────────────────────────────────────────────────────────

function doGet(e) {
  var data = {
    committees: getCommitteeNamesForForm(),
    targets:    getUpcomingCouncilTargets()
  };
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Client sends: Content-Type: text/plain; body is JSON string.
function doPost(e) {
  try {
    var data        = JSON.parse(e.postData.contents);
    var committee   = String(data.committee     || '').trim();
    var authorName  = String(data.authorName    || '').trim();
    var authorEmail = String(data.authorEmail   || '').trim();
    var target      = String(data.councilTarget || '').trim();
    var bodyHTML    = String(data.bodyHTML      || '').trim();

    if (!committee || !authorName || !authorEmail || !target || !bodyHTML) {
      throw new Error('All fields are required.');
    }
    if (bodyHTML === '<p><br></p>') {
      throw new Error('Report body cannot be empty.');
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('reports');
    if (!sheet) throw new Error('"reports" tab not found in the spreadsheet.');

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([now, committee, authorName, authorEmail, target, bodyHTML]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helpers used by MeetingPacket.js ─────────────────────────────────────────

// Returns committee names (skipping ~ officer rows) for the form dropdown.
function getCommitteeNamesForForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_committees');
  var data  = sheet.getDataRange().getValues();
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || '');
    if (!name || name.charAt(0) === '~') continue;
    names.push(name);
  }
  return names;
}

// Returns upcoming Council meeting months in "MMMM yyyy" format (e.g. "June 2026"), sorted.
function getUpcomingCouncilTargets() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var data  = ss.getSheetByName('leadership_meetings').getDataRange().getValues();
  var tz    = Session.getScriptTimeZone();
  var today = new Date();
  var seen  = {}, targets = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] !== 'Council') continue;
    var d = data[i][0];
    if (!(d instanceof Date) || d < today) continue;
    var label = Utilities.formatDate(d, tz, 'MMMM yyyy');
    if (!seen[label]) { seen[label] = true; targets.push({ label: label, time: d.getTime() }); }
  }
  targets.sort(function(a, b) { return a.time - b.time; });
  return targets.map(function(t) { return t.label; });
}

// Returns all submitted reports whose councilTarget matches `target` (e.g. "June 2026").
function getCommitteeReportsForTarget(target) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('reports');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() !== target) continue;
    result.push({
      submittedDate: data[i][0],
      committee:     String(data[i][1]),
      authorName:    String(data[i][2]),
      authorEmail:   String(data[i][3]),
      councilTarget: String(data[i][4]),
      bodyHTML:      String(data[i][5])
    });
  }
  return result;
}

// ── Quill HTML → Google Doc elements ─────────────────────────────────────────
//
// Inserts Quill-generated HTML into a Google Doc body starting at startIdx.
// Preserves bold, italic, links, headings (h1–h3), bullet lists, and ordered lists.
// Returns the next available insertion index.

function insertQuillHtmlIntoDoc(body, html, startIdx) {
  var blocks = parseQuillBlocks(html);
  var idx = startIdx;
  var headingMap = {
    heading1: DocumentApp.ParagraphHeading.HEADING1,
    heading2: DocumentApp.ParagraphHeading.HEADING2,
    heading3: DocumentApp.ParagraphHeading.HEADING3
  };
  blocks.forEach(function(block) {
    try {
      if (headingMap[block.type]) {
        var para = body.insertParagraph(idx, block.text);
        para.setHeading(headingMap[block.type]);
        applyInlineFormatting(para.editAsText(), block.runs);
        idx++;
      } else if (block.type === 'bullet' || block.type === 'ordered') {
        var item = body.insertListItem(idx, block.text);
        var attrs = {};
        attrs[DocumentApp.Attribute.NESTING_LEVEL] = 0;
        attrs[DocumentApp.Attribute.GLYPH_TYPE] = block.type === 'bullet'
          ? DocumentApp.GlyphType.BULLET
          : DocumentApp.GlyphType.DECIMAL;
        item.setAttributes(attrs);
        applyInlineFormatting(item.editAsText(), block.runs);
        idx++;
      } else {
        var para = body.insertParagraph(idx, block.text);
        applyInlineFormatting(para.editAsText(), block.runs);
        idx++;
      }
    } catch (e) { /* skip elements that can't be inserted cross-document */ }
  });
  return idx;
}

// Parses top-level block elements from Quill HTML into a flat block list.
function parseQuillBlocks(html) {
  var blocks = [];
  html = html.replace(/>\s+</g, '><').trim();
  var pos = 0;

  while (pos < html.length) {
    if (html[pos] !== '<') { pos++; continue; }
    var m = html.slice(pos).match(/^<(h[1-6]|p|ul|ol)(?:\s[^>]*)?>/i);
    if (!m) { pos++; continue; }
    var tagName = m[1].toLowerCase();
    pos += m[0].length;

    // Find matching close tag, tracking nesting depth.
    var closeTag = '</' + tagName + '>';
    var depth = 1, searchPos = pos, closePos = -1;
    while (depth > 0 && searchPos < html.length) {
      var openRe  = new RegExp('<' + tagName + '[\\s>]', 'i');
      var openOff = html.slice(searchPos).search(openRe);
      var openIdx = openOff === -1 ? -1 : openOff + searchPos;
      var closeIdx = html.toLowerCase().indexOf(closeTag, searchPos);
      if (closeIdx === -1) { searchPos = html.length; break; }
      if (openIdx !== -1 && openIdx < closeIdx) {
        depth++; searchPos = openIdx + tagName.length + 1;
      } else {
        depth--; if (depth === 0) closePos = closeIdx;
        searchPos = closeIdx + closeTag.length;
      }
    }
    if (closePos === -1) continue;

    var content = html.substring(pos, closePos);
    pos = closePos + closeTag.length;

    if (tagName === 'ul' || tagName === 'ol') {
      var listType = tagName === 'ul' ? 'bullet' : 'ordered';
      var liRe = /<li(?:[^>]*)>([\s\S]*?)<\/li>/gi;
      var liM;
      while ((liM = liRe.exec(content)) !== null) {
        var runs = parseInlineRuns(liM[1]);
        var text = runsToText(runs);
        if (text.trim()) blocks.push({ type: listType, runs: runs, text: text });
      }
    } else {
      var type = tagName === 'p' ? 'paragraph'
        : tagName === 'h1' ? 'heading1'
        : tagName === 'h2' ? 'heading2'
        : 'heading3';
      var runs = parseInlineRuns(content);
      var text = runsToText(runs);
      if (text.trim()) blocks.push({ type: type, runs: runs, text: text });
    }
  }
  return blocks;
}

// Parses inline HTML (strong, em, a, br) into styled text runs.
function parseInlineRuns(html) {
  var runs  = [];
  var stack = [{ bold: false, italic: false, link: null }];
  var pos   = 0;
  var cur   = '';
  html = html.replace(/<br\s*\/?>/gi, '\n');

  while (pos < html.length) {
    if (html[pos] !== '<') { cur += html[pos++]; continue; }
    var end = html.indexOf('>', pos);
    if (end === -1) { cur += html[pos++]; continue; }
    var fullTag  = html.substring(pos, end + 1);
    var inner    = fullTag.slice(1, -1).trim();
    var closing  = inner[0] === '/';
    if (closing) inner = inner.slice(1);
    var spIdx    = inner.search(/[\s\/]/);
    var tagName  = (spIdx === -1 ? inner : inner.slice(0, spIdx)).toLowerCase();

    if (cur) {
      var s = stack[stack.length - 1];
      runs.push({ text: decodeHtmlEntities(cur), bold: s.bold, italic: s.italic, link: s.link });
      cur = '';
    }
    pos = end + 1;

    if (!closing) {
      var s = Object.assign({}, stack[stack.length - 1]);
      if (tagName === 'strong' || tagName === 'b') s.bold = true;
      else if (tagName === 'em' || tagName === 'i') s.italic = true;
      else if (tagName === 'a') {
        var hm = fullTag.match(/href="([^"]*)"/i);
        s.link = hm ? hm[1] : null;
      }
      if (fullTag.slice(-2) !== '/>') stack.push(s);
    } else {
      if (stack.length > 1) stack.pop();
    }
  }
  if (cur) {
    var s = stack[stack.length - 1];
    runs.push({ text: decodeHtmlEntities(cur), bold: s.bold, italic: s.italic, link: s.link });
  }
  return runs;
}

function runsToText(runs) {
  return runs.map(function(r) { return r.text; }).join('');
}

function decodeHtmlEntities(text) {
  return text
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, function(m, n) { return String.fromCharCode(parseInt(n, 10)); });
}

function applyInlineFormatting(textEl, runs) {
  var pos = 0;
  runs.forEach(function(run) {
    if (!run.text) return;
    var end = pos + run.text.length - 1;
    if (end >= pos) {
      if (run.bold)   textEl.setBold(pos, end, true);
      if (run.italic) textEl.setItalic(pos, end, true);
      if (run.link)   textEl.setLinkUrl(pos, end, run.link);
    }
    pos += run.text.length;
  });
}
