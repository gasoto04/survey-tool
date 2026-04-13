import { useState } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";

var COLORS = ["#003366","#0063B2","#4A90D9","#89B4E8","#C4D9F0","#F4A261","#E76F51","#2A9D8F","#264653","#E9C46A","#7C3AED","#DB2777","#059669","#D97706","#6366F1"];
var LK = ["To a great extent","To a moderate extent","To a slight extent","Not at all","Don't know or not applicable"];
var LC = {"To a great extent":"#003366","To a moderate extent":"#4A90D9","To a slight extent":"#89B4E8","Not at all":"#E76F51","Don't know or not applicable":"#C4D9F0"};
// Extended Likert patterns for different surveys
var LK_PATTERNS = [
  ["To a great extent","To a moderate extent","To a slight extent","Not at all"],
  ["Strongly agree","Agree","Disagree","Strongly disagree"],
  ["Always","Often","Sometimes","Rarely","Never"],
  ["Extensive guidance","Some guidance","No guidance","Not applicable"],
  ["Very well","Well","Fairly well","Not well","Not at all"],
  ["Yes, to a great extent","Yes, to some extent","No","Don't know"]
];

function detectLikertScale(values) {
  var nonEmpty = values.filter(function(v) { return v !== null && v !== undefined && String(v).trim() !== ""; });
  if (nonEmpty.length === 0) return null;
  // Check standard LK
  if (nonEmpty.some(function(v) { return LK.indexOf(String(v).trim()) > -1; })) return LK;
  // Check extended patterns
  for (var p = 0; p < LK_PATTERNS.length; p++) {
    var pattern = LK_PATTERNS[p];
    var match = nonEmpty.some(function(v) { return pattern.indexOf(String(v).trim()) > -1; });
    if (match) {
      // Get all unique values that belong to this pattern + any extras
      var unique = {};
      nonEmpty.forEach(function(v) { unique[String(v).trim()] = true; });
      var allInPattern = Object.keys(unique).every(function(u) {
        return pattern.indexOf(u) > -1 || u.toLowerCase().indexOf("don't know") > -1 || u.toLowerCase().indexOf("not applicable") > -1 || u.toLowerCase() === "n/a";
      });
      if (allInPattern) {
        var scale = pattern.slice();
        Object.keys(unique).forEach(function(u) { if (scale.indexOf(u) === -1) scale.push(u); });
        return scale;
      }
    }
  }
  return null;
}

function classifyQuestion(values) {
  var nonEmpty = values.filter(function(v) { return v !== null && v !== undefined && String(v).trim() !== ""; });
  if (nonEmpty.length === 0) return "empty";
  var unique = {};
  var totalLen = 0;
  nonEmpty.forEach(function(v) { var s = String(v).trim(); unique[s] = (unique[s] || 0) + 1; totalLen += s.length; });
  var uniqueCount = Object.keys(unique).length;
  var avgLen = totalLen / nonEmpty.length;
  // Check for any Likert pattern
  if (detectLikertScale(nonEmpty)) return "likert";
  var lows = Object.keys(unique).map(function(k){return k.toLowerCase();});
  if (lows.every(function(k){return k==="yes"||k==="no";})) return "categorical";
  if (uniqueCount <= 10 && avgLen < 80) return "categorical";
  if (uniqueCount > nonEmpty.length * 0.4 && avgLen > 30) return "openended";
  if (avgLen > 100) return "openended";
  if (uniqueCount <= 20) return "categorical";
  return "openended";
}

function getSamples(values, n) {
  var nonEmpty = values.filter(function(v) { return v !== null && v !== undefined && String(v).trim() !== ""; });
  var unique = []; var seen = {};
  nonEmpty.forEach(function(v) { var s = String(v).trim(); if (!seen[s] && unique.length < n) { seen[s] = true; unique.push(s); } });
  return unique;
}

function getValueCounts(values) {
  var counts = {};
  values.forEach(function(v) { if (v !== null && v !== undefined && String(v).trim() !== "") { var s = String(v).trim(); counts[s] = (counts[s] || 0) + 1; } });
  return counts;
}

// Get the Likert scale for a question (returns LK default or detected scale)
function getLikertScale(q, data) {
  if (q.detectedScale) return q.detectedScale;
  return LK;
}

// Generate Likert color map for any scale
function getLikertColors(scale) {
  var baseColors = ["#003366","#4A90D9","#89B4E8","#E76F51","#C4D9F0","#F4A261","#E9C46A","#94a3b8"];
  var colors = {};
  scale.forEach(function(s, i) { colors[s] = baseColors[i % baseColors.length]; });
  return colors;
}

export default function App() {
  var _data = useState(null), data = _data[0], setData = _data[1];
  var _raw = useState(null), rawData = _raw[0], setRawData = _raw[1];
  var _headers = useState([]), headers = _headers[0], setHeaders = _headers[1];
  var _ci = useState(null), cleanInfo = _ci[0], setCleanInfo = _ci[1];
  var _dupes = useState([]), dupeRows = _dupes[0], setDupeRows = _dupes[1];
  var _dupeChecked = useState({}), dupeChecked = _dupeChecked[0], setDupeChecked = _dupeChecked[1];
  var _v = useState("upload"), view = _v[0], setView = _v[1];
  var _p = useState({}), picks = _p[0], setPicks = _p[1];
  var _ct = useState({}), cTypes = _ct[0], setCTypes = _ct[1];
  var _nr = useState({}), narr = _nr[0], setNarr = _nr[1];
  var _nl = useState({}), narrLoading = _nl[0], setNarrLoading = _nl[1];
  var _os = useState({}), openSec = _os[0], setOpenSec = _os[1];
  var _ps = useState(null), parsed = _ps[0], setParsed = _ps[1];
  var _prev = useState({}), showPreview = _prev[0], setShowPreview = _prev[1];
  var _chartPrev = useState({}), showChartPreview = _chartPrev[0], setShowChartPreview = _chartPrev[1];
  var _expandQ = useState({}), expandQ = _expandQ[0], setExpandQ = _expandQ[1];
  var _expStatus = useState(""), expStatus = _expStatus[0], setExpStatus = _expStatus[1];
  var _editPrompt = useState({}), editPrompt = _editPrompt[0], setEditPrompt = _editPrompt[1];
  var _showEdit = useState({}), showEdit = _showEdit[0], setShowEdit = _showEdit[1];
  var _exportHTML = useState(null), exportHTML = _exportHTML[0], setExportHTML = _exportHTML[1];
  var _copyStatus = useState({}), copyStatus = _copyStatus[0], setCopyStatus = _copyStatus[1];

  // ── UNIVERSAL FILE PARSER ──
  function onFile(e) {
    var file = e.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function(evt) {
      try {
        var wb = XLSX.read(evt.target.result, { type: "array" });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        if (rows.length < 2) { alert("File has no data rows."); return; }
        var h = rows[0].map(function(x) { return x ? String(x) : ""; });
        var d = rows.slice(1).filter(function(row) {
          // Remove completely empty rows
          return row.some(function(v) { return v !== null && v !== undefined && String(v).trim() !== ""; });
        });
        if (d.length === 0) { alert("No data rows found."); return; }
        setHeaders(h);
        setRawData(d);

        // Smart duplicate detection: find email/identifier columns
        var emailIdx = -1; var countryIdx = -1; var nameIdx = -1; var lnIdx = -1;
        h.forEach(function(hdr, i) {
          var low = hdr.toLowerCase();
          if (low.indexOf("email") > -1 && emailIdx === -1) emailIdx = i;
          if ((low.indexOf("country") > -1 || low.indexOf("(merged)") > -1) && low.indexOf("select") > -1) countryIdx = i;
          if (low === "first name") nameIdx = i;
          if (low === "last name") lnIdx = i;
        });
        // If no email, try response code for unique ID
        var idIdx = emailIdx;
        if (idIdx < 0) {
          h.forEach(function(hdr, i) {
            if (hdr.toLowerCase().indexOf("response code") > -1 && idIdx < 0) idIdx = i;
          });
        }

        var seen = {};
        var foundDupes = [];
        var autoChecked = {};
        d.forEach(function(row, i) {
          var id = idIdx >= 0 ? String(row[idIdx] || "").toLowerCase().trim() : "";
          var co = countryIdx >= 0 ? String(row[countryIdx] || "") : "";
          if (!id) return; // Can't detect dupes without identifier
          var key = id + "|" + co;
          if (seen[key] !== undefined) {
            foundDupes.push({
              idx: i,
              email: emailIdx >= 0 ? String(row[emailIdx] || "") : (idIdx >= 0 ? String(row[idIdx] || "") : ""),
              name: (nameIdx >= 0 ? String(row[nameIdx] || "") : "") + " " + (lnIdx >= 0 ? String(row[lnIdx] || "") : ""),
              country: co,
              origIdx: seen[key],
              row: row
            });
            autoChecked[i] = true;
          } else {
            seen[key] = i;
          }
        });
        setDupeRows(foundDupes);
        setDupeChecked(autoChecked);
        setCleanInfo({ orig: d.length, dupeCount: foundDupes.length });
        setView("clean");
      } catch (err) { alert("Error reading file: " + err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function confirmClean() {
    var toRemove = {};
    dupeRows.forEach(function(d) { if (dupeChecked[d.idx]) toRemove[d.idx] = true; });
    var cleaned = rawData.filter(function(_, i) { return !toRemove[i]; });
    setData(cleaned);

    // ── UNIVERSAL SECTION & QUESTION PARSER ──
    var secs = {};

    // Step 1: Identify metadata columns and section headers
    var metaCols = {};
    var sectionPositions = []; // [{idx, name, letter}]

    headers.forEach(function(hdr, i) {
      if (!hdr || hdr.trim() === "") { metaCols[i] = true; return; }
      var low = hdr.toLowerCase().trim();
      // Known metadata patterns
      if (low === "response code" || low === "first name" || low === "last name" ||
          low.indexOf("email") > -1 || low.indexOf("completed date") > -1 ||
          low === "questions" || low === "question" ||
          low.indexOf("response id") > -1 || low.indexOf("start date") > -1 ||
          low.indexOf("end date") > -1 || low.indexOf("ip address") > -1 ||
          low.indexOf("collector") > -1 || low === "status" || low === "custom data 1") {
        metaCols[i] = true;
        return;
      }
      // Check if column has ALL empty data → likely a section header
      var hasData = cleaned.some(function(row) {
        return row[i] !== null && row[i] !== undefined && String(row[i]).trim() !== "";
      });
      if (!hasData) {
        var sm = hdr.match(/^([A-Z])\.\s+(.+)/);
        sectionPositions.push({ idx: i, name: sm ? sm[2].trim() : hdr.trim(), letter: sm ? sm[1] : null });
        metaCols[i] = true;
        return;
      }
      // Inline "A. Section" that also has data
      var sm2 = hdr.match(/^([A-Z])\.\s+(.+)/);
      if (sm2 && !hdr.match(/^\d/)) {
        sectionPositions.push({ idx: i, name: sm2[2].trim(), letter: sm2[1] });
      }
    });

    // Default section if none found
    if (sectionPositions.length === 0) {
      sectionPositions.push({ idx: -1, name: "Survey Questions", letter: "A" });
    }

    // Assign unique keys to sections
    var usedLetters = {};
    sectionPositions.forEach(function(sp, si) {
      if (sp.letter) {
        sp.key = sp.letter;
      } else {
        // Auto-assign: S1, S2, etc.
        sp.key = "S" + (si + 1);
      }
      usedLetters[sp.key] = true;
    });

    // Step 2: Merged columns
    var mergedColIdx = {};
    var skipCols = {};
    headers.forEach(function(hdr, i) {
      if (hdr && hdr.indexOf("(Merged)") > -1) {
        var baseText = hdr.replace(/\s*\(Merged\)\s*/, "").trim().toLowerCase();
        mergedColIdx[baseText] = i;
      }
    });
    Object.keys(mergedColIdx).forEach(function(baseText) {
      headers.forEach(function(hdr, i) {
        if (!hdr || i === mergedColIdx[baseText]) return;
        var cleanedHdr = hdr.replace(/^\d+\.?\d*\.?\s*/, "").trim().toLowerCase();
        if (cleanedHdr === baseText || cleanedHdr.indexOf(baseText) === 0) {
          skipCols[i] = true;
        }
      });
    });

    // Step 3: Helper functions
    function getSectionForCol(colIdx) {
      var sec = sectionPositions[0];
      for (var s = 0; s < sectionPositions.length; s++) {
        if (sectionPositions[s].idx <= colIdx) sec = sectionPositions[s];
        else break;
      }
      return sec;
    }

    function extractSubItem(headerText) {
      var cleanHdr = headerText.replace(/\s*\(Merged\)\s*/, "").trim();
      // Pattern 1: "::" separator (e.g. "question:: * sub-item")
      if (cleanHdr.indexOf("::") > -1) {
        var parts = cleanHdr.split("::");
        var main = parts[0].trim();
        var sub = parts.slice(1).join("::").trim().replace(/^\*\s*/, "").trim();
        return { main: main, sub: sub || null };
      }
      // Pattern 2: "question?: * sub" or "question on?: * sub"
      var m2 = cleanHdr.match(/^(.+\?)\s*:\s*\*\s*(.+)$/);
      if (m2) return { main: m2[1].trim(), sub: m2[2].trim() };
      // Pattern 3: "question: * sub" (non-greedy, sub must be < 200 chars)
      var m3 = cleanHdr.match(/^(.{20,}):\s*\*\s*(.{1,200})$/);
      if (m3) return { main: m3[1].trim(), sub: m3[2].trim() };
      return { main: cleanHdr, sub: null };
    }

    function extractQNum(text) {
      var dm = text.match(/^(\d+\.?\d*\.?)\s+\d+\.?\d*\.?\s/);
      if (dm) return dm[1].replace(/\.+$/, "");
      var sm = text.match(/^(\d+\.?\d*\.?)\s/);
      if (sm) return sm[1].replace(/\.+$/, "");
      var qm = text.match(/^Q(\d+\.?\d*)/i);
      if (qm) return qm[1];
      return "";
    }

    function cleanMainText(text) {
      var c = text.replace(/^\d+\.?\d*\.?\s+\d+\.?\d*\.?\s*/, "");
      if (c === text) c = text.replace(/^\d+\.?\d*\.?\s*/, "");
      return c.trim();
    }

    // Step 4: Parse each column into questions
    var qCounter = 0;
    headers.forEach(function(hdr, i) {
      if (!hdr || metaCols[i] || skipCols[i]) return;
      var hasData = cleaned.some(function(row) {
        return row[i] !== null && row[i] !== undefined && String(row[i]).trim() !== "";
      });
      if (!hasData) return;

      var sec = getSectionForCol(i);
      var secKey = sec.key;
      var extracted = extractSubItem(hdr);
      var qNum = extractQNum(extracted.main);
      var mainClean = cleanMainText(extracted.main);
      var sub = extracted.sub;

      // Fallback: if mainClean is empty, use the full header
      if (!mainClean || mainClean.length < 3) mainClean = hdr.trim();

      if (!secs[secKey]) secs[secKey] = { name: sec.name, qs: {} };

      if (!secs[secKey].qs[mainClean]) {
        qCounter++;
        var colValues = cleaned.map(function(row) { return row[i]; });
        var qType = classifyQuestion(colValues);
        var detectedScale = detectLikertScale(colValues);
        secs[secKey].qs[mainClean] = {
          main: mainClean, qNum: qNum || String(qCounter), subs: [], cols: [],
          qType: qType, samples: getSamples(colValues, 6), detectedScale: detectedScale
        };
      } else {
        var existQ = secs[secKey].qs[mainClean];
        var newS = getSamples(cleaned.map(function(row) { return row[i]; }), 3);
        newS.forEach(function(s) { if (existQ.samples.indexOf(s) === -1 && existQ.samples.length < 8) existQ.samples.push(s); });
        if (existQ.qType !== "likert") {
          var cv = cleaned.map(function(row) { return row[i]; });
          var ns = detectLikertScale(cv);
          if (ns) { existQ.qType = "likert"; existQ.detectedScale = ns; }
        }
      }
      secs[secKey].qs[mainClean].subs.push(sub || mainClean);
      secs[secKey].qs[mainClean].cols.push(i);
    });

    // Remove empty sections
    Object.keys(secs).forEach(function(sk) {
      if (Object.keys(secs[sk].qs).length === 0) delete secs[sk];
    });

    if (Object.keys(secs).length === 0) {
      alert("Could not detect any survey questions. Please ensure the file has headers in the first row and data in subsequent rows.");
      return;
    }

    setParsed(secs);
    var autoP = {};
    var autoC = {};
    Object.keys(secs).forEach(function(sk) {
      Object.keys(secs[sk].qs).forEach(function(qk) {
        var k = sk + "|" + qk;
        var q = secs[sk].qs[qk];
        autoP[k] = true;
        if (q.qType === "openended") autoC[k] = "text";
        else if (q.qType === "likert" && q.subs.length > 1) autoC[k] = "stacked";
        else if (q.qType === "categorical") autoC[k] = "bar";
        else if (q.qType === "likert" && q.subs.length === 1) autoC[k] = "pie";
        else if (q.subs.length === 1) autoC[k] = "pie";
        else autoC[k] = "stacked";
      });
    });
    setPicks(autoP);
    setCTypes(autoC);
    setCleanInfo(function(prev) { return Object.assign({}, prev, { cleaned: cleaned.length, removed: Object.keys(toRemove).length }); });
    setView("config");
  }

  // ── LOAD html2canvas DYNAMICALLY ──
  var _imgOverlay = useState(null), imgOverlay = _imgOverlay[0], setImgOverlay = _imgOverlay[1];

  function loadH2C() {
    return new Promise(function(resolve, reject) {
      if (window.html2canvas) { resolve(window.html2canvas); return; }
      var s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
      s.onload = function() { resolve(window.html2canvas); };
      s.onerror = function() { reject(new Error("Failed to load html2canvas")); };
      document.head.appendChild(s);
    });
  }

  // ── SCREENSHOT CHART → show as image overlay (right-click to save) ──
  function screenshotChart(chartId, qNum) {
    var el = document.getElementById(chartId);
    if (!el) return;
    setCopyStatus(function(p) { return Object.assign({}, p, {[chartId]: "copying"}); });
    loadH2C().then(function(html2canvas) {
      return html2canvas(el, { backgroundColor: "#fafbfc", scale: 2, useCORS: true, logging: false });
    }).then(function(canvas) {
      var dataUrl = canvas.toDataURL("image/png");
      setImgOverlay({ src: dataUrl, label: "Q" + qNum });
      setCopyStatus(function(p) { return Object.assign({}, p, {[chartId]: "done"}); });
    }).catch(function(err) {
      setCopyStatus(function(p) { return Object.assign({}, p, {[chartId]: "fail"}); });
      alert("Screenshot failed: " + err.message);
    });
    setTimeout(function() { setCopyStatus(function(p) { return Object.assign({}, p, {[chartId]: ""}); }); }, 3000);
  }

  // ── AI NARRATIVE ──
  function doNarrative(sk, sName) {
    if (!data || !parsed) return;
    setNarrLoading(function(p) { return Object.assign({}, p, {[sk]: true}); });
    var sec = parsed[sk];
    var lines = Object.keys(sec.qs).map(function(qk) {
      var q = sec.qs[qk];
      if (q.qType === "openended") {
        var samples = getSamples(getQValues(q), 6);
        return "Q" + q.qNum + " (open-ended): " + q.main + "\nSample responses: " + samples.join(" | ");
      }
      var vals = getQValues(q);
      return "Q" + q.qNum + ": " + q.main + "\nResponses: " + JSON.stringify(getValueCounts(vals));
    }).join("\n\n");

    fetch("https://api.anthropic.com/v1/messages", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514", max_tokens: 1000,
        messages: [{ role: "user", content: "You are an IMF IEO analyst. Write ONE paragraph of 4 sentences summarizing key findings from survey section \"" + sName + "\". Include percentages for quantitative questions and themes for open-ended. Professional tone. No bullets.\n\nIMPORTANT: After each factual claim, add an inline reference in square brackets. ALWAYS include the question number — format: [Q3: response, 45%] or [Q13.1: response, 12%]. NEVER write just [Q: ...] without a number.\n\n" + lines }]
      })
    }).then(function(r) { return r.json(); }).then(function(d) {
      var text = d.content && d.content[0] ? d.content[0].text : "Failed.";
      setNarr(function(p) { return Object.assign({}, p, {[sk]: text}); });
      setNarrLoading(function(p) { return Object.assign({}, p, {[sk]: false}); });
    }).catch(function(err) {
      setNarr(function(p) { return Object.assign({}, p, {[sk]: "Error: " + err.message}); });
      setNarrLoading(function(p) { return Object.assign({}, p, {[sk]: false}); });
    });
  }

  function refineNarrative(sk, sName) {
    var instructions = editPrompt[sk];
    if (!instructions || !instructions.trim()) return;
    setNarrLoading(function(p) { return Object.assign({}, p, {[sk]: true}); });
    var sec = parsed[sk];
    var dataLines = Object.keys(sec.qs).map(function(qk) {
      var q = sec.qs[qk];
      return "Q" + q.qNum + ": " + q.main + "\nResponses: " + JSON.stringify(getValueCounts(getQValues(q)));
    }).join("\n\n");

    fetch("https://api.anthropic.com/v1/messages", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514", max_tokens: 1000,
        messages: [{ role: "user", content: "You are an IMF IEO analyst. You previously wrote this narrative for survey section \"" + sName + "\":\n\n\"" + narr[sk] + "\"\n\nThe underlying data:\n" + dataLines + "\n\nThe user wants the following changes:\n" + instructions + "\n\nRewrite incorporating these changes. ONE paragraph, 4 sentences. Professional tone. No bullets.\n\nIMPORTANT: After each factual claim, add [Q3: response, 45%] references. ALWAYS include question number.\n\nOutput only the revised paragraph." }]
      })
    }).then(function(r) { return r.json(); }).then(function(d) {
      var text = d.content && d.content[0] ? d.content[0].text : "Failed.";
      setNarr(function(p) { return Object.assign({}, p, {[sk]: text}); });
      setNarrLoading(function(p) { return Object.assign({}, p, {[sk]: false}); });
      setEditPrompt(function(p) { return Object.assign({}, p, {[sk]: ""}); });
    }).catch(function(err) {
      setNarr(function(p) { return Object.assign({}, p, {[sk]: "Error: " + err.message}); });
      setNarrLoading(function(p) { return Object.assign({}, p, {[sk]: false}); });
    });
  }

  // ── GET VALUES ──
  function getQValues(q) {
    if (!data) return [];
    if (q.combinedValues && q.combinedValues.length > 0) return q.combinedValues;
    var vals = [];
    q.cols.forEach(function(ci) { data.forEach(function(row) { if (row[ci] !== null && row[ci] !== undefined && String(row[ci]).trim() !== "") vals.push(String(row[ci]).trim()); }); });
    return vals;
  }

  // ── EXPORT: render inline report ──
  function exportWord() {
    if (!parsed || !data) return;
    var html = '';
    html += '<h1 style="color:#003366;font-size:18pt;border-bottom:2px solid #003366;padding-bottom:6px;font-family:Calibri,sans-serif;">Survey Analysis Report</h1>';
    html += '<p style="color:#64748b;font-family:Calibri,sans-serif;font-size:11pt;">Generated ' + new Date().toLocaleDateString() + ' | ' + data.length + ' responses</p>';

    Object.keys(parsed).forEach(function(sk) {
      var sec = parsed[sk];
      var activeQs = Object.keys(sec.qs).filter(function(qk) { return picks[sk + "|" + qk]; });
      if (activeQs.length === 0) return;
      html += '<h2 style="color:#003366;font-size:14pt;margin-top:24pt;font-family:Calibri,sans-serif;">Section ' + sk + ': ' + sec.name + '</h2>';
      if (narr[sk]) html += '<div style="background:#f0fdfa;border-left:4px solid #2A9D8F;padding:12px 16px;margin:12pt 0;line-height:1.6;font-family:Calibri,sans-serif;font-size:11pt;"><b>AI Narrative:</b><br/>' + narr[sk] + '</div>';

      activeQs.forEach(function(qk) {
        var q = sec.qs[qk];
        html += '<h3 style="color:#0063B2;font-size:12pt;font-family:Calibri,sans-serif;">Q' + q.qNum + ' ' + q.main + '</h3>';
        var vals = getQValues(q);
        var scale = getLikertScale(q, data);

        if (cTypes[sk+"|"+qk] === "text" || q.qType === "openended") {
          var samples = getSamples(vals, 10);
          html += '<p style="color:#64748b;font-size:10pt;">Open-ended · ' + vals.length + ' responses</p>';
          samples.forEach(function(s) { html += '<div style="background:#f8fafc;padding:6px 10px;margin:4pt 0;border-left:3px solid #89B4E8;font-style:italic;font-family:Calibri,sans-serif;font-size:11pt;">"' + s + '"</div>'; });
        } else if (q.qType === "likert" && q.subs.length > 1) {
          html += '<table style="border-collapse:collapse;width:100%;margin:10pt 0;font-family:Calibri,sans-serif;"><tr>';
          html += '<th style="background:#003366;color:white;padding:6px 8px;font-size:10pt;text-align:left;border:1px solid #003366;">Sub-question</th>';
          scale.forEach(function(l) { html += '<th style="background:#003366;color:white;padding:6px 8px;font-size:10pt;text-align:center;border:1px solid #003366;">' + l + '</th>'; });
          html += '<th style="background:#003366;color:white;padding:6px 8px;font-size:10pt;text-align:center;border:1px solid #003366;">n</th></tr>';
          q.subs.forEach(function(sq, si) {
            var counts = {}; scale.forEach(function(l){counts[l]=0;});
            data.forEach(function(row) { var v = row[q.cols[si]]; if (v) { var vs = String(v).trim(); if (counts[vs] !== undefined) counts[vs]++; } });
            var total = scale.reduce(function(s,l){return s+counts[l];},0);
            html += '<tr><td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:left;">' + sq + '</td>';
            scale.forEach(function(l) { html += '<td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:center;">' + counts[l] + ' (' + (total?Math.round(counts[l]/total*100):0) + '%)</td>'; });
            html += '<td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:center;font-weight:bold;">' + total + '</td></tr>';
          });
          html += '</table>';
        } else {
          var counts = getValueCounts(vals);
          var total = vals.length;
          var sorted = Object.keys(counts).sort(function(a,b){return counts[b]-counts[a];});
          html += '<table style="border-collapse:collapse;width:100%;margin:10pt 0;font-family:Calibri,sans-serif;">';
          html += '<tr><th style="background:#003366;color:white;padding:6px 8px;font-size:10pt;text-align:left;border:1px solid #003366;">Response</th>';
          html += '<th style="background:#003366;color:white;padding:6px 8px;font-size:10pt;text-align:center;border:1px solid #003366;">Count</th>';
          html += '<th style="background:#003366;color:white;padding:6px 8px;font-size:10pt;text-align:center;border:1px solid #003366;">%</th></tr>';
          sorted.forEach(function(k) { html += '<tr><td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:left;">' + k + '</td><td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:center;">' + counts[k] + '</td><td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:center;">' + Math.round(counts[k]/total*100) + '%</td></tr>'; });
          html += '<tr><td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:left;font-weight:bold;">Total</td><td style="border:1px solid #ddd;padding:5px 8px;font-size:10pt;text-align:center;font-weight:bold;">' + total + '</td><td></td></tr></table>';
        }
      });
    });
    setExportHTML(html);
  }

  function selectAllReport() {
    var el = document.getElementById("export-report-content");
    if (!el) return;
    var range = document.createRange();
    range.selectNodeContents(el);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    setExpStatus("✓ Selected! Now press Ctrl+C (or Cmd+C) to copy, then paste into Word.");
    setTimeout(function() { setExpStatus(""); }, 6000);
  }

  // ── RENDER NARRATIVE WITH PARSED REFERENCES ──
  function renderNarrativeWithRefs(text) {
    if (!text) return null;
    var parts = [];
    var regex = /\[(Q[\d.]*)\s*:\s*([^\]]+)\]/g;
    var lastIdx = 0; var match;
    while ((match = regex.exec(text)) !== null) {
      if (match.index > lastIdx) parts.push({ type: "text", content: text.substring(lastIdx, match.index) });
      parts.push({ type: "ref", label: match[1], detail: match[2].trim() });
      lastIdx = match.index + match[0].length;
    }
    if (lastIdx < text.length) parts.push({ type: "text", content: text.substring(lastIdx) });
    if (parts.length === 0) return text;
    return parts.map(function(p, i) {
      if (p.type === "ref") {
        return <span key={i} title={"Reference " + p.label + " — " + p.detail}
          style={{ display: "inline", background: "#dbeafe", color: "#1e40af", padding: "1px 6px", borderRadius: 3, fontSize: 11, fontWeight: 600, cursor: "help", whiteSpace: "nowrap", marginLeft: 2, marginRight: 2, borderBottom: "2px solid #93c5fd" }}>
          {p.label.length > 1 ? p.label : p.label + ": " + (p.detail.length > 25 ? p.detail.substring(0, 22) + "..." : p.detail)}
        </span>;
      }
      return <span key={i}>{p.content}</span>;
    });
  }

  // ── MINI CHART PREVIEW (config screen) ──
  function renderMiniChart(q, k) {
    if (!data) return null;
    var type = cTypes[k] || "stacked";
    var scale = getLikertScale(q, data);
    var lc = getLikertColors(scale);

    if (type === "text") {
      var vals = getQValues(q);
      var unique = getSamples(vals, 5);
      return <div style={{ maxHeight: 150, overflowY: "auto" }}>
        {unique.map(function(v, i) {
          return <div key={i} style={{ padding: "4px 8px", margin: "2px 0", background: "white", borderLeft: "2px solid #89B4E8", fontSize: 11, color: "#334155" }}>"{v.length > 80 ? v.substring(0, 77) + "..." : v}"</div>;
        })}
      </div>;
    }
    if (type === "freqtable") {
      var fvals = getQValues(q);
      var fcounts = getValueCounts(fvals);
      var ftotal = fvals.length;
      var fsorted = Object.keys(fcounts).sort(function(a,b){return fcounts[b]-fcounts[a];}).slice(0, 6);
      return <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
        <thead><tr style={{ background: "#003366", color: "white" }}>
          <th style={{ padding: "3px 6px", textAlign: "left" }}>Response</th>
          <th style={{ padding: "3px 6px", textAlign: "center", width: 45 }}>Count</th>
          <th style={{ padding: "3px 6px", textAlign: "center", width: 40 }}>%</th>
        </tr></thead>
        <tbody>{fsorted.map(function(name, i) {
          return <tr key={i} style={{ background: i % 2 === 0 ? "#f8fafc" : "white" }}>
            <td style={{ padding: "2px 6px" }}>{name.length > 35 ? name.substring(0, 32) + "..." : name}</td>
            <td style={{ padding: "2px 6px", textAlign: "center" }}>{fcounts[name]}</td>
            <td style={{ padding: "2px 6px", textAlign: "center" }}>{Math.round(fcounts[name] / ftotal * 100)}%</td>
          </tr>;
        })}</tbody>
      </table>;
    }
    if (type === "pie") {
      var pvals = getQValues(q);
      var pcounts = getValueCounts(pvals);
      var pd = Object.keys(pcounts).map(function(n){return {name:n,value:pcounts[n]};}).sort(function(a,b){return b.value-a.value;});
      var ptotal = pvals.length;
      return <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <ResponsiveContainer width="45%" height={140}>
          <PieChart><Pie data={pd} cx="50%" cy="50%" outerRadius={55} innerRadius={22} dataKey="value"
            label={function(e){return Math.round(e.percent*100)+"%";}}>
            {pd.map(function(_,i){return <Cell key={i} fill={COLORS[i%COLORS.length]} />;})}
          </Pie></PieChart>
        </ResponsiveContainer>
        <div style={{ fontSize: 10 }}>
          {pd.slice(0, 5).map(function(d, i) {
            return <div key={i} style={{ display: "flex", alignItems: "center", gap: 4, marginBottom: 2 }}>
              <div style={{ width: 8, height: 8, borderRadius: 2, background: COLORS[i % COLORS.length], flexShrink: 0 }} />
              <span>{d.name.length > 20 ? d.name.substring(0, 17) + "..." : d.name}: {Math.round(d.value / ptotal * 100)}%</span>
            </div>;
          })}
        </div>
      </div>;
    }
    if (type === "bar") {
      var bvals = getQValues(q);
      var bcounts = getValueCounts(bvals);
      var bdata = Object.keys(bcounts).sort(function(a,b){return bcounts[b]-bcounts[a];}).slice(0, 8).map(function(n) {
        return { name: n.length > 18 ? n.substring(0, 15) + "..." : n, count: bcounts[n] };
      });
      return <ResponsiveContainer width="100%" height={Math.max(120, bdata.length * 22)}>
        <BarChart data={bdata} layout="vertical" margin={{ left: 90, right: 10, top: 2, bottom: 2 }}>
          <XAxis type="number" tick={{ fontSize: 9 }} />
          <YAxis dataKey="name" type="category" width={88} tick={{ fontSize: 9 }} />
          <Bar dataKey="count" fill="#003366" radius={[0, 3, 3, 0]}>
            {bdata.map(function(_, i) { return <Cell key={i} fill={COLORS[i % COLORS.length]} />; })}
          </Bar>
        </BarChart>
      </ResponsiveContainer>;
    }
    // STACKED — works for any number of subs
    if (type === "stacked") {
      if (q.subs.length > 1) {
        var sdata = q.subs.slice(0, 5).map(function(sq, i) {
          var counts = {}; scale.forEach(function(l) { counts[l] = 0; });
          data.forEach(function(row) { var v = row[q.cols[i]]; if (v) { var vs = String(v).trim(); if (counts[vs] !== undefined) counts[vs]++; } });
          var total = scale.reduce(function(s, l) { return s + counts[l]; }, 0);
          var r = { name: sq.length > 25 ? sq.substring(0, 22) + "..." : sq };
          scale.forEach(function(l) { r[l] = total > 0 ? Math.round(counts[l] / total * 100) : 0; });
          return r;
        });
        return <ResponsiveContainer width="100%" height={Math.max(100, sdata.length * 28)}>
          <BarChart data={sdata} layout="vertical" margin={{ left: 100, right: 10, top: 2, bottom: 2 }}>
            <XAxis type="number" domain={[0, 100]} tickFormatter={function(v){return v+"%";}} tick={{ fontSize: 9 }} />
            <YAxis dataKey="name" type="category" width={98} tick={{ fontSize: 8 }} />
            {scale.map(function(l) { return <Bar key={l} dataKey={l} stackId="a" fill={lc[l] || "#ccc"} />; })}
          </BarChart>
        </ResponsiveContainer>;
      } else {
        // Single sub stacked → render as bar chart
        var sv = getQValues(q);
        var sc = getValueCounts(sv);
        var sd = Object.keys(sc).sort(function(a,b){return sc[b]-sc[a];}).slice(0, 8).map(function(n) {
          return { name: n.length > 18 ? n.substring(0, 15) + "..." : n, count: sc[n] };
        });
        return <ResponsiveContainer width="100%" height={Math.max(100, sd.length * 22)}>
          <BarChart data={sd} layout="vertical" margin={{ left: 90, right: 10, top: 2, bottom: 2 }}>
            <XAxis type="number" tick={{ fontSize: 9 }} />
            <YAxis dataKey="name" type="category" width={88} tick={{ fontSize: 9 }} />
            <Bar dataKey="count" fill="#003366" radius={[0, 3, 3, 0]}>
              {sd.map(function(_, i) { return <Cell key={i} fill={COLORS[i % COLORS.length]} />; })}
            </Bar>
          </BarChart>
        </ResponsiveContainer>;
      }
    }
    if (type === "table") {
      var trows = q.subs.slice(0, 4).map(function(sq, i) {
        var counts = {}; scale.forEach(function(l) { counts[l] = 0; });
        data.forEach(function(row) { var v = row[q.cols[i]]; if (v) { var vs = String(v).trim(); if (counts[vs] !== undefined) counts[vs]++; } });
        var total = scale.reduce(function(s, l) { return s + counts[l]; }, 0);
        return { label: sq.length > 30 ? sq.substring(0, 27) + "..." : sq, counts: counts, total: total };
      });
      return <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 9 }}>
        <thead><tr style={{ background: "#003366", color: "white" }}>
          <th style={{ padding: "2px 4px", textAlign: "left" }}>Sub-Q</th>
          {scale.map(function(l) { return <th key={l} style={{ padding: "2px 3px", textAlign: "center" }}>{l.split(" ").slice(0, 2).join(" ").substring(0, 12)}</th>; })}
        </tr></thead>
        <tbody>{trows.map(function(r, i) {
          return <tr key={i} style={{ background: i % 2 === 0 ? "#f8fafc" : "white" }}>
            <td style={{ padding: "2px 4px" }}>{r.label}</td>
            {scale.map(function(l) { return <td key={l} style={{ padding: "2px 3px", textAlign: "center" }}>{r.counts[l]}</td>; })}
          </tr>;
        })}</tbody>
      </table>;
    }
    return <div style={{ fontSize: 11, color: "#94a3b8", fontStyle: "italic" }}>No preview for this chart type.</div>;
  }

  // ── FULL CHART RENDERS ──
  function chartOptions(q, k) {
    var type = cTypes[k] || "stacked";
    var opts = [];
    if (q.qType === "likert" && q.subs.length > 1) {
      opts = [{val:"stacked",icon:"📊",label:"Stacked"},{val:"table",icon:"📋",label:"Table"},{val:"text",icon:"📝",label:"Text"}];
    } else if (q.qType === "categorical" || (q.qType !== "openended" && q.qType !== "likert")) {
      opts = [{val:"bar",icon:"📊",label:"Bar"},{val:"pie",icon:"🥧",label:"Pie"},{val:"freqtable",icon:"📋",label:"Freq Table"},{val:"text",icon:"📝",label:"Text"}];
    } else if (q.qType === "openended") {
      opts = [{val:"text",icon:"📝",label:"Text"},{val:"freqtable",icon:"📋",label:"Freq Table"}];
    } else {
      opts = [{val:"stacked",icon:"📊",label:"Stacked"},{val:"pie",icon:"🥧",label:"Pie"},{val:"bar",icon:"📊",label:"Bar"},{val:"freqtable",icon:"📋",label:"Table"},{val:"text",icon:"📝",label:"Text"}];
    }
    return <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10, flexWrap: "wrap" }}>
      <span style={{ fontSize: 11, color: "#64748b" }}>Chart:</span>
      {opts.map(function(o) {
        return <button key={o.val} onClick={function() { setCTypes(function(p) { return Object.assign({}, p, {[k]: o.val}); }); }}
          style={{ padding: "3px 9px", borderRadius: 4, border: type === o.val ? "2px solid #003366" : "1px solid #ddd", background: type === o.val ? "#e8f0fe" : "white", fontSize: 11, cursor: "pointer", fontWeight: type === o.val ? 600 : 400, color: type === o.val ? "#003366" : "#666" }}>
          {o.icon} {o.label}
        </button>;
      })}
    </div>;
  }

  function renderChart(q, k) {
    var type = cTypes[k] || "stacked";
    var sel = chartOptions(q, k);
    var chartId = "chart-" + k.replace(/[^a-zA-Z0-9]/g, "-");
    var copyBtn = <button onClick={function() { screenshotChart(chartId, q.qNum); }}
      title="Screenshot this chart — right-click the image to save"
      style={{ padding: "3px 10px", borderRadius: 4, border: "1px solid #cbd5e1", background: copyStatus[chartId] === "done" ? "#d1fae5" : "#f8fafc", fontSize: 10, cursor: "pointer", color: copyStatus[chartId] === "done" ? "#065f46" : "#64748b", fontWeight: 500, flexShrink: 0 }}>
      {copyStatus[chartId] === "done" ? "✓ Done!" : copyStatus[chartId] === "copying" ? "📸..." : "📸 Screenshot"}
    </button>;

    if (type === "text") return renderText(q, k, sel, chartId, copyBtn);
    if (type === "freqtable") return renderFreqTable(q, k, sel, chartId, copyBtn);
    if (type === "pie") return renderPie(q, k, sel, chartId, copyBtn);
    if (type === "bar") return renderBar(q, k, sel, chartId, copyBtn);
    if (type === "table") return renderLikertTable(q, k, sel, chartId, copyBtn);
    return renderStacked(q, k, sel, chartId, copyBtn);
  }

  function renderStacked(q, k, sel, chartId, copyBtn) {
    var scale = getLikertScale(q, data);
    var lc = getLikertColors(scale);
    var chartData = q.subs.map(function(sq, i) {
      var counts = {}; scale.forEach(function(l) { counts[l] = 0; });
      data.forEach(function(row) { var v = row[q.cols[i]]; if (v) { var vs = String(v).trim(); if (counts[vs] !== undefined) counts[vs]++; } });
      var total = scale.reduce(function(s, l) { return s + counts[l]; }, 0);
      var r = { name: sq.length > 35 ? sq.substring(0, 32) + "..." : sq };
      scale.forEach(function(l) { r[l] = total > 0 ? Math.round(counts[l] / total * 100) : 0; });
      return r;
    });
    return <div id={chartId} key={k} style={{ marginBottom: 24, padding: 16, background: "#fafbfc", borderRadius: 8, border: "1px solid #eef1f5" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <h4 style={{ fontSize: 13, color: "#003366", margin: "0 0 6px", flex: 1 }}>Q{q.qNum} {q.main}</h4>
        {copyBtn}
      </div>
      {sel}
      <ResponsiveContainer width="100%" height={Math.max(200, q.subs.length * 42)}>
        <BarChart data={chartData} layout="vertical" margin={{ left: 150, right: 20, top: 5, bottom: 5 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
          <XAxis type="number" domain={[0, 100]} tickFormatter={function(v){return v+"%";}} tick={{fontSize:11}} />
          <YAxis dataKey="name" type="category" width={145} tick={{fontSize:10}} />
          <Tooltip formatter={function(v){return v+"%";}} />
          <Legend wrapperStyle={{fontSize:10}} />
          {scale.map(function(l) { return <Bar key={l} dataKey={l} stackId="a" fill={lc[l] || "#ccc"} />; })}
        </BarChart>
      </ResponsiveContainer>
    </div>;
  }

  function renderBar(q, k, sel, chartId, copyBtn) {
    var vals = getQValues(q);
    var counts = getValueCounts(vals);
    var total = vals.length;
    var chartData = Object.keys(counts).sort(function(a,b){return counts[b]-counts[a];}).map(function(name) {
      return { name: name.length > 25 ? name.substring(0, 22) + "..." : name, count: counts[name], pct: Math.round(counts[name]/total*100) };
    });
    return <div id={chartId} key={k} style={{ marginBottom: 24, padding: 16, background: "#fafbfc", borderRadius: 8, border: "1px solid #eef1f5" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <h4 style={{ fontSize: 13, color: "#003366", margin: "0 0 6px", flex: 1 }}>Q{q.qNum} {q.main}</h4>
        {copyBtn}
      </div>
      <span style={{ fontSize: 11, color: "#94a3b8" }}>n = {total}</span>
      {sel}
      <ResponsiveContainer width="100%" height={Math.max(200, chartData.length * 32)}>
        <BarChart data={chartData} layout="vertical" margin={{ left: 140, right: 30, top: 5, bottom: 5 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
          <XAxis type="number" tick={{fontSize:11}} />
          <YAxis dataKey="name" type="category" width={135} tick={{fontSize:10}} />
          <Tooltip formatter={function(v,n){return n==="pct"?v+"%":v;}} />
          <Bar dataKey="count" fill="#003366" radius={[0,4,4,0]}>
            {chartData.map(function(_,i){return <Cell key={i} fill={COLORS[i%COLORS.length]} />;})}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>;
  }

  function renderPie(q, k, sel, chartId, copyBtn) {
    var vals = getQValues(q);
    var counts = getValueCounts(vals);
    var pd = Object.keys(counts).map(function(n){return {name:n, value:counts[n]};}).sort(function(a,b){return b.value-a.value;});
    var total = vals.length;
    return <div id={chartId} key={k} style={{ marginBottom: 24, padding: 16, background: "#fafbfc", borderRadius: 8, border: "1px solid #eef1f5" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <h4 style={{ fontSize: 13, color: "#003366", margin: "0 0 6px", flex: 1 }}>Q{q.qNum} {q.main}</h4>
        {copyBtn}
      </div>
      {sel}
      <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
        <ResponsiveContainer width="45%" height={220}>
          <PieChart><Pie data={pd} cx="50%" cy="50%" outerRadius={85} innerRadius={35} dataKey="value"
            label={function(e){return Math.round(e.percent*100)+"%";}}>
            {pd.map(function(_,i){return <Cell key={i} fill={COLORS[i%COLORS.length]} />;})}
          </Pie><Tooltip /></PieChart>
        </ResponsiveContainer>
        <div style={{fontSize:12, maxHeight:200, overflowY:"auto"}}>
          {pd.map(function(d,i){
            return <div key={i} style={{display:"flex",alignItems:"center",gap:6,marginBottom:3}}>
              <div style={{width:10,height:10,borderRadius:2,background:COLORS[i%COLORS.length],flexShrink:0}} />
              <span>{d.name}: <b>{d.value}</b> ({Math.round(d.value/total*100)}%)</span>
            </div>;
          })}
          <div style={{marginTop:4,color:"#999",fontStyle:"italic"}}>n={total}</div>
        </div>
      </div>
    </div>;
  }

  function renderFreqTable(q, k, sel, chartId, copyBtn) {
    var vals = getQValues(q);
    var counts = getValueCounts(vals);
    var total = vals.length;
    var sorted = Object.keys(counts).sort(function(a,b){return counts[b]-counts[a];});
    return <div id={chartId} key={k} style={{ marginBottom: 24, padding: 16, background: "#fafbfc", borderRadius: 8, border: "1px solid #eef1f5" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <h4 style={{ fontSize: 13, color: "#003366", margin: "0 0 6px", flex: 1 }}>Q{q.qNum} {q.main}</h4>
        {copyBtn}
      </div>
      {sel}
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
          <thead><tr style={{ background: "#003366", color: "white" }}>
            <th style={{ padding: "6px 10px", textAlign: "left" }}>Response</th>
            <th style={{ padding: "6px 8px", textAlign: "center", width: 70 }}>Count</th>
            <th style={{ padding: "6px 8px", textAlign: "center", width: 70 }}>%</th>
            <th style={{ padding: "6px 8px", textAlign: "left", width: "30%" }}>Bar</th>
          </tr></thead>
          <tbody>
            {sorted.map(function(name, i) {
              var pct = Math.round(counts[name] / total * 100);
              return <tr key={i} style={{ background: i % 2 === 0 ? "#f8fafc" : "white", borderBottom: "1px solid #e8e8e8" }}>
                <td style={{ padding: "5px 10px", textAlign: "left" }}>{name}</td>
                <td style={{ padding: "5px 8px", textAlign: "center" }}>{counts[name]}</td>
                <td style={{ padding: "5px 8px", textAlign: "center" }}>{pct}%</td>
                <td style={{ padding: "5px 8px" }}>
                  <div style={{ background: "#e2e8f0", borderRadius: 3, height: 16, width: "100%" }}>
                    <div style={{ background: COLORS[i % COLORS.length], borderRadius: 3, height: 16, width: pct + "%", minWidth: pct > 0 ? 4 : 0 }} />
                  </div>
                </td>
              </tr>;
            })}
            <tr style={{ borderTop: "2px solid #003366", fontWeight: 600 }}>
              <td style={{ padding: "5px 10px", textAlign: "left" }}>Total</td>
              <td style={{ padding: "5px 8px", textAlign: "center" }}>{total}</td>
              <td style={{ padding: "5px 8px", textAlign: "center" }}>100%</td>
              <td></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>;
  }

  function renderLikertTable(q, k, sel, chartId, copyBtn) {
    var scale = getLikertScale(q, data);
    var rows = q.subs.map(function(sq, i) {
      var counts = {}; scale.forEach(function(l){counts[l]=0;});
      data.forEach(function(row) { var v = row[q.cols[i]]; if (v) { var vs = String(v).trim(); if (counts[vs] !== undefined) counts[vs]++; } });
      var total = scale.reduce(function(s,l){return s+counts[l];},0);
      return { label: sq, counts: counts, total: total };
    });
    return <div id={chartId} key={k} style={{ marginBottom: 24, padding: 16, background: "#fafbfc", borderRadius: 8, border: "1px solid #eef1f5" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <h4 style={{ fontSize: 13, color: "#003366", margin: "0 0 6px", flex: 1 }}>Q{q.qNum} {q.main}</h4>
        {copyBtn}
      </div>
      {sel}
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
          <thead><tr style={{ background: "#003366", color: "white" }}>
            <th style={{ padding: "6px 8px", textAlign: "left" }}>Sub-question</th>
            {scale.map(function(l) { return <th key={l} style={{ padding: "6px 4px", textAlign: "center" }}>{l}</th>; })}
            <th style={{ padding: "6px 4px", textAlign: "center" }}>n</th>
          </tr></thead>
          <tbody>{rows.map(function(r, i) {
            return <tr key={i} style={{ background: i % 2 === 0 ? "#f8fafc" : "white", borderBottom: "1px solid #e8e8e8" }}>
              <td style={{ padding: "5px 8px", textAlign: "left" }}>{r.label}</td>
              {scale.map(function(l) { return <td key={l} style={{ padding: "5px 4px", textAlign: "center" }}>{r.counts[l]} <span style={{ color: "#bbb", fontSize: 9 }}>({r.total ? Math.round(r.counts[l] / r.total * 100) : 0}%)</span></td>; })}
              <td style={{ padding: "5px 4px", textAlign: "center", fontWeight: 600 }}>{r.total}</td>
            </tr>;
          })}</tbody>
        </table>
      </div>
    </div>;
  }

  function renderText(q, k, sel, chartId, copyBtn) {
    var vals = getQValues(q);
    var unique = getSamples(vals, 15);
    return <div id={chartId} key={k} style={{ marginBottom: 24, padding: 16, background: "#fafbfc", borderRadius: 8, border: "1px solid #eef1f5" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <h4 style={{ fontSize: 13, color: "#003366", margin: "0 0 6px", flex: 1 }}>Q{q.qNum} {q.main}</h4>
        {copyBtn}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
        <span style={{ background: "#fef3c7", color: "#92400e", padding: "2px 8px", borderRadius: 4, fontSize: 10, fontWeight: 600 }}>
          {q.qType === "openended" ? "OPEN-ENDED" : q.qType.toUpperCase()}
        </span>
        <span style={{ fontSize: 11, color: "#64748b" }}>{vals.length} responses · {unique.length} unique</span>
      </div>
      {sel}
      <div style={{ maxHeight: 280, overflowY: "auto" }}>
        {unique.map(function(v, i) {
          return <div key={i} style={{ padding: "7px 12px", margin: "3px 0", background: "white", borderLeft: "3px solid #89B4E8", borderRadius: "0 5px 5px 0", fontSize: 12, lineHeight: 1.5, color: "#334155" }}>{v}</div>;
        })}
        {vals.length > 15 && <div style={{ fontSize: 11, color: "#94a3b8", padding: "6px 0", fontStyle: "italic" }}>...and {vals.length - 15} more</div>}
      </div>
    </div>;
  }

  // ── STYLES ──
  var box = { background: "white", borderRadius: 10, padding: 22, margin: "14px 20px", boxShadow: "0 1px 8px rgba(0,0,0,0.06)", border: "1px solid #e2e8f0" };
  var tagStyle = function(t) {
    var c = { likert: {bg:"#dbeafe",c:"#1e40af"}, categorical: {bg:"#d1fae5",c:"#065f46"}, openended: {bg:"#fef3c7",c:"#92400e"}, empty: {bg:"#f1f5f9",c:"#94a3b8"} };
    var col = c[t] || c.empty;
    return { background: col.bg, color: col.c, padding: "2px 7px", borderRadius: 4, fontSize: 10, fontWeight: 600, display: "inline-block" };
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f0f4f8", fontFamily: "system-ui,sans-serif" }}>
      <div style={{ background: "linear-gradient(135deg,#001d3d,#003566)", padding: "18px 24px", color: "white", display: "flex", alignItems: "center" }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 9, letterSpacing: 3, textTransform: "uppercase", opacity: 0.5 }}>IEO Survey Tool</div>
          <div style={{ fontSize: 19, fontWeight: 700 }}>Survey Analysis & Visualization</div>
          {cleanInfo && cleanInfo.cleaned && <div style={{ fontSize: 12, opacity: 0.7, marginTop: 2 }}>{cleanInfo.cleaned} responses</div>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {expStatus && <span style={{fontSize:11,color:"#E9C46A"}}>{expStatus}</span>}
          {view !== "upload" && (
            <button onClick={function() {
              if (view === "viz") setView("config");
              else if (view === "config") setView("clean");
              else if (view === "clean") { setView("upload"); setData(null); setParsed(null); setNarr({}); setCleanInfo(null); setRawData(null); setDupeRows([]); setDupeChecked({}); setExportHTML(null); setCopyStatus({}); }
            }} style={{ padding: "6px 14px", background: "rgba(255,255,255,0.15)", color: "white", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 5, fontSize: 11, fontWeight: 500, cursor: "pointer" }}>
              ← Back
            </button>
          )}
          {view === "viz" && (
            <button onClick={exportWord} style={{ padding: "6px 14px", background: "#2A9D8F", color: "white", border: "none", borderRadius: 5, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>📄 Export</button>
          )}
          {view !== "upload" && (
            <button onClick={function() { setView("upload"); setData(null); setParsed(null); setNarr({}); setCleanInfo(null); setRawData(null); setDupeRows([]); setDupeChecked({}); setExportHTML(null); setCopyStatus({}); }}
              style={{ padding: "6px 14px", background: "rgba(255,255,255,0.1)", color: "white", border: "1px solid rgba(255,255,255,0.25)", borderRadius: 5, fontSize: 11, fontWeight: 500, cursor: "pointer" }}>
              📁 New File
            </button>
          )}
        </div>
      </div>

      {/* IMAGE OVERLAY — for chart screenshots */}
      {imgOverlay && (
        <div onClick={function() { setImgOverlay(null); }}
          style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, zIndex: 10000, background: "rgba(0,0,0,0.7)", display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", cursor: "pointer" }}>
          <div style={{ background: "white", borderRadius: 10, padding: 16, maxWidth: "90%", maxHeight: "90vh", overflow: "auto", boxShadow: "0 8px 32px rgba(0,0,0,0.3)" }}
            onClick={function(e) { e.stopPropagation(); }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: "#003366" }}>📸 {imgOverlay.label} — Right-click image → "Save image as..."</div>
              <button onClick={function() { setImgOverlay(null); }}
                style={{ padding: "2px 8px", border: "1px solid #ccc", borderRadius: 4, background: "#f8fafc", fontSize: 12, cursor: "pointer" }}>✕</button>
            </div>
            <img src={imgOverlay.src} alt={imgOverlay.label} style={{ maxWidth: "100%", border: "1px solid #e2e8f0", borderRadius: 6 }} />
            <div style={{ marginTop: 8, fontSize: 11, color: "#64748b", textAlign: "center" }}>
              💡 Right-click (or long-press) the image → <b>Save image as...</b> or <b>Copy image</b>
            </div>
          </div>
        </div>
      )}

      {/* EXPORT MODAL — inline rendered report with Select All */}
      {exportHTML && (
        <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, zIndex: 9999, background: "rgba(0,0,0,0.6)", display: "flex", justifyContent: "center", alignItems: "center" }}>
          <div style={{ background: "white", borderRadius: 12, width: "90%", maxWidth: 850, maxHeight: "90vh", display: "flex", flexDirection: "column", overflow: "hidden", boxShadow: "0 8px 32px rgba(0,0,0,0.3)" }}>
            <div style={{ padding: "12px 20px", background: "#003366", color: "white", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
              <div style={{ fontWeight: 600, fontSize: 14 }}>📄 Export Report</div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                {expStatus && <span style={{ fontSize: 11, color: "#E9C46A" }}>{expStatus}</span>}
                <button onClick={selectAllReport}
                  style={{ padding: "5px 14px", background: "#2A9D8F", color: "white", border: "none", borderRadius: 4, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
                  ✅ Select All → then Ctrl+C
                </button>
                <button onClick={function() { setExportHTML(null); setExpStatus(""); }}
                  style={{ padding: "5px 10px", background: "transparent", color: "white", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 4, fontSize: 12, cursor: "pointer" }}>✕</button>
              </div>
            </div>
            <div style={{ flex: 1, overflow: "auto", padding: "20px 28px", background: "white" }}>
              <div id="export-report-content" dangerouslySetInnerHTML={{ __html: exportHTML }} />
            </div>
            <div style={{ padding: "8px 20px", background: "#f0f4f8", borderTop: "1px solid #e2e8f0", fontSize: 11, color: "#64748b", flexShrink: 0 }}>
              💡 Click <b>"Select All"</b> above, then <b>Ctrl+C</b> (Cmd+C on Mac), then paste into Word or Google Docs. Tables and formatting will be preserved.
            </div>
          </div>
        </div>
      )}

      {/* UPLOAD */}
      {view === "upload" && (
        <div style={Object.assign({}, box, { textAlign: "center", padding: "50px 20px" })}>
          <div style={{ fontSize: 42, marginBottom: 10 }}>📊</div>
          <div style={{ fontSize: 18, fontWeight: 600, color: "#003366", marginBottom: 6 }}>Upload Survey Excel File</div>
          <div style={{ fontSize: 13, color: "#64748b", marginBottom: 6 }}>Auto-detects question types, finds duplicates for review, generates charts & AI narratives</div>
          <div style={{ fontSize: 11, color: "#94a3b8", marginBottom: 22 }}>Supports any sheet name · Multiple question formats · .xlsx, .xls, .csv</div>
          <label style={{ display: "inline-block", padding: "11px 28px", background: "#003366", color: "white", borderRadius: 7, cursor: "pointer", fontSize: 14, fontWeight: 600 }}>
            📁 Choose File
            <input type="file" accept=".xlsx,.xls,.csv" onChange={onFile} style={{ display: "none" }} />
          </label>
        </div>
      )}

      {/* CLEAN — duplicate review */}
      {view === "clean" && cleanInfo && (
        <div style={box}>
          <div style={{ fontSize: 15, fontWeight: 600, color: "#003366", marginBottom: 2 }}>🧹 Data Cleaning — Review Duplicates</div>
          <div style={{ fontSize: 10, color: "#94a3b8", marginBottom: 12 }}>Step 1 of 3</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
            <div style={{ background: "#f8fafc", borderRadius: 6, padding: "10px 14px", borderLeft: "4px solid #64748b", minWidth: 110 }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: "#64748b" }}>{cleanInfo.orig}</div>
              <div style={{ fontSize: 11, color: "#64748b" }}>Total Rows</div>
            </div>
            <div style={{ background: "#f8fafc", borderRadius: 6, padding: "10px 14px", borderLeft: "4px solid #E76F51", minWidth: 110 }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: "#E76F51" }}>{dupeRows.length}</div>
              <div style={{ fontSize: 11, color: "#64748b" }}>Duplicates Found</div>
            </div>
            <div style={{ background: "#f8fafc", borderRadius: 6, padding: "10px 14px", borderLeft: "4px solid #003366", minWidth: 110 }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: "#003366" }}>{cleanInfo.orig - Object.keys(dupeChecked).filter(function(k){return dupeChecked[k];}).length}</div>
              <div style={{ fontSize: 11, color: "#64748b" }}>After Removal</div>
            </div>
          </div>

          {dupeRows.length > 0 ? (
            <div>
              <div style={{ fontSize: 12, color: "#64748b", marginBottom: 8 }}>Review duplicate rows below. Uncheck any you want to <b>keep</b>.</div>
              <div style={{ marginBottom: 8, display: "flex", gap: 8 }}>
                <button onClick={function() { var all = {}; dupeRows.forEach(function(d){all[d.idx]=true;}); setDupeChecked(all); }}
                  style={{ padding: "4px 12px", fontSize: 11, borderRadius: 4, border: "1px solid #ccc", background: "#f8fafc", cursor: "pointer" }}>Select All</button>
                <button onClick={function() { setDupeChecked({}); }}
                  style={{ padding: "4px 12px", fontSize: 11, borderRadius: 4, border: "1px solid #ccc", background: "#f8fafc", cursor: "pointer" }}>Deselect All</button>
              </div>
              <div style={{ maxHeight: 320, overflowY: "auto", border: "1px solid #e2e8f0", borderRadius: 6 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead><tr style={{ background: "#003366", color: "white", position: "sticky", top: 0 }}>
                    <th style={{ padding: "6px 8px", width: 36 }}>✓</th>
                    <th style={{ padding: "6px 8px", textAlign: "left" }}>Row</th>
                    <th style={{ padding: "6px 8px", textAlign: "left" }}>Identifier</th>
                    <th style={{ padding: "6px 8px", textAlign: "left" }}>Country</th>
                    <th style={{ padding: "6px 8px", textAlign: "left" }}>Dup of Row</th>
                  </tr></thead>
                  <tbody>
                    {dupeRows.map(function(d, i) {
                      return <tr key={i} style={{ background: dupeChecked[d.idx] ? "#fef2f2" : i % 2 === 0 ? "#f8fafc" : "white" }}>
                        <td style={{ padding: "5px 8px", textAlign: "center" }}>
                          <input type="checkbox" checked={!!dupeChecked[d.idx]} onChange={function() { setDupeChecked(function(p) { var n = Object.assign({}, p); if (n[d.idx]) delete n[d.idx]; else n[d.idx] = true; return n; }); }} />
                        </td>
                        <td style={{ padding: "5px 8px" }}>{d.idx + 2}</td>
                        <td style={{ padding: "5px 8px", color: "#4A90D9" }}>{d.email || d.name.trim()}</td>
                        <td style={{ padding: "5px 8px" }}>{d.country || "—"}</td>
                        <td style={{ padding: "5px 8px", color: "#94a3b8" }}>Row {d.origIdx + 2}</td>
                      </tr>;
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div style={{ padding: "12px 16px", background: "#d1fae5", borderRadius: 6, fontSize: 13, color: "#065f46" }}>✅ No duplicates found! Data is clean.</div>
          )}

          <button onClick={confirmClean} style={{ marginTop: 14, padding: "10px 28px", background: "#003366", color: "white", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
            {dupeRows.length > 0 ? "Confirm & Remove " + Object.keys(dupeChecked).filter(function(k){return dupeChecked[k];}).length + " Rows →" : "Continue →"}
          </button>
        </div>
      )}

      {/* CONFIGURE */}
      {view === "config" && parsed && data && (
        <div style={box}>
          <div style={{ fontSize: 15, fontWeight: 600, color: "#003366", marginBottom: 4 }}>⚙️ Configure Visualizations</div>
          <div style={{ fontSize: 10, color: "#94a3b8", marginBottom: 2 }}>Step 2 of 3 · <button onClick={function(){ setView("clean"); }} style={{ border: "none", background: "none", color: "#4A90D9", fontSize: 10, cursor: "pointer", textDecoration: "underline", padding: 0 }}>← Back to Cleaning</button></div>
          <div style={{ fontSize: 12, color: "#64748b", marginBottom: 14 }}>Questions auto-classified with recommended chart types (★). Use <b>📋 Responses</b> for sample data, <b>📊 Chart Preview</b> to preview.</div>
          {Object.keys(parsed).map(function(sk) {
            var sec = parsed[sk];
            return <div key={sk} style={{ marginBottom: 14 }}>
              <div onClick={function() { setOpenSec(function(p) { return Object.assign({}, p, {[sk]: !p[sk]}); }); }}
                style={{ padding: "8px 14px", background: "#003366", color: "white", borderRadius: openSec[sk] ? "6px 6px 0 0" : 6, fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", justifyContent: "space-between" }}>
                <span>Section {sk}: {sec.name}</span><span>{openSec[sk] ? "▾" : "▸"} ({Object.keys(sec.qs).length} questions)</span>
              </div>
              {openSec[sk] && <div style={{ border: "1px solid #e2e8f0", borderTop: "none", borderRadius: "0 0 6px 6px", padding: 10 }}>
                {Object.keys(sec.qs).map(function(qk) {
                  var q = sec.qs[qk];
                  var k = sk + "|" + qk;
                  var optionsList = [];
                  if (q.qType === "likert" && q.subs.length > 1) optionsList = [{v:"stacked",l:"Stacked Bar"},{v:"table",l:"Likert Table"},{v:"text",l:"Text"}];
                  else if (q.qType === "likert" && q.subs.length === 1) optionsList = [{v:"pie",l:"Pie"},{v:"bar",l:"Bar"},{v:"freqtable",l:"Freq Table"},{v:"text",l:"Text"}];
                  else if (q.qType === "categorical") optionsList = [{v:"bar",l:"Bar Chart"},{v:"pie",l:"Pie"},{v:"freqtable",l:"Freq Table"},{v:"text",l:"Text"}];
                  else if (q.qType === "openended") optionsList = [{v:"text",l:"Text"},{v:"freqtable",l:"Freq Table"}];
                  else optionsList = [{v:"stacked",l:"Stacked"},{v:"bar",l:"Bar"},{v:"pie",l:"Pie"},{v:"freqtable",l:"Table"},{v:"text",l:"Text"}];
                  var isExpanded = expandQ[k];
                  var qText = q.main;
                  var isLong = qText.length > 75;

                  return <div key={k} style={{ marginBottom: 4, padding: "10px 8px", borderBottom: "1px solid #f1f5f9", background: picks[k] ? "white" : "#fafafa" }}>
                    <div style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 12 }}>
                      <input type="checkbox" checked={!!picks[k]} onChange={function() { setPicks(function(p) { return Object.assign({}, p, {[k]: !p[k]}); }); }} style={{ marginTop: 3 }} />
                      <span style={tagStyle(q.qType)}>{({likert:"Likert",categorical:"Categorical",openended:"Open-ended",empty:"Empty"})[q.qType]}</span>
                      <div style={{ flex: 1, color: "#1e293b", lineHeight: 1.5 }}>
                        <b style={{ color: "#003366" }}>Q{q.qNum}</b>{" "}
                        {isLong && !isExpanded ? qText.substring(0, 75) + "..." : qText}
                        {q.subs.length > 1 && <span style={{ color: "#94a3b8" }}> ({q.subs.length} sub-items)</span>}
                        {isLong && (
                          <button onClick={function() { setExpandQ(function(p) { return Object.assign({}, p, {[k]: !p[k]}); }); }}
                            style={{ marginLeft: 4, padding: "0 4px", border: "none", background: "none", color: "#4A90D9", fontSize: 11, cursor: "pointer", textDecoration: "underline" }}>
                            {isExpanded ? "less" : "more"}
                          </button>
                        )}
                      </div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 6, marginLeft: 24 }}>
                      <span style={{ fontSize: 10, color: "#94a3b8" }}>Chart:</span>
                      <select value={cTypes[k] || "stacked"} onChange={function(e) { var v = e.target.value; setCTypes(function(p) { return Object.assign({}, p, {[k]: v}); }); setShowChartPreview(function(p) { return Object.assign({}, p, {[k]: true}); }); }}
                        style={{ padding: "3px 6px", borderRadius: 4, border: "1px solid #ccc", fontSize: 11, background: "white" }}>
                        {optionsList.map(function(o) { return <option key={o.v} value={o.v}>{o.l + (o.v === optionsList[0].v ? " ★" : "")}</option>; })}
                      </select>
                      <span style={{ fontSize: 9, color: "#94a3b8" }}>★ recommended</span>
                      <div style={{ flex: 1 }} />
                      <button onClick={function() { setShowPreview(function(p) { return Object.assign({}, p, {[k]: !p[k]}); }); setShowChartPreview(function(p) { return Object.assign({}, p, {[k]: false}); }); }}
                        style={{ padding: "3px 10px", borderRadius: 4, border: showPreview[k] ? "2px solid #003366" : "1px solid #ddd", background: showPreview[k] ? "#e8f0fe" : "#f8fafc", fontSize: 10, cursor: "pointer", color: "#003366", fontWeight: showPreview[k] ? 600 : 400 }}>
                        {showPreview[k] ? "📋 Hide" : "📋 Responses"}
                      </button>
                      <button onClick={function() { setShowChartPreview(function(p) { return Object.assign({}, p, {[k]: !p[k]}); }); setShowPreview(function(p) { return Object.assign({}, p, {[k]: false}); }); }}
                        style={{ padding: "3px 10px", borderRadius: 4, border: showChartPreview[k] ? "2px solid #0A9396" : "1px solid #ddd", background: showChartPreview[k] ? "#d1fae5" : "#f8fafc", fontSize: 10, cursor: "pointer", color: "#065f46", fontWeight: showChartPreview[k] ? 600 : 400 }}>
                        {showChartPreview[k] ? "📊 Hide" : "📊 Chart Preview"}
                      </button>
                    </div>
                    {showPreview[k] && (
                      <div style={{ marginTop: 8, marginLeft: 24, padding: "10px 12px", background: "#f8fafc", borderRadius: 6, border: "1px solid #e8ecf0" }}>
                        <div style={{ fontSize: 10, color: "#64748b", marginBottom: 4, fontWeight: 600 }}>Sample responses ({q.samples.length}):</div>
                        {q.samples.map(function(s, i) {
                          return <div key={i} style={{ fontSize: 11, color: "#334155", padding: "3px 0", borderBottom: i < q.samples.length - 1 ? "1px solid #eee" : "none" }}>
                            {s.length > 150 ? s.substring(0, 147) + "..." : s}
                          </div>;
                        })}
                        {q.subs.length > 1 && <div style={{ marginTop: 6, fontSize: 10, color: "#94a3b8" }}>Sub-items: {q.subs.slice(0, 3).map(function(s){return s.length > 50 ? s.substring(0, 47) + "..." : s;}).join(" · ")}{q.subs.length > 3 ? " · +" + (q.subs.length - 3) + " more" : ""}</div>}
                      </div>
                    )}
                    {showChartPreview[k] && (
                      <div style={{ marginTop: 8, marginLeft: 24, padding: "12px", background: "#f0fdf4", borderRadius: 6, border: "1px solid #bbf7d0" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                          <div style={{ fontSize: 10, color: "#065f46", fontWeight: 600 }}>Preview — {(cTypes[k] || "stacked").replace("freqtable","Frequency Table").replace("stacked","Stacked Bar").replace("bar","Bar Chart").replace("pie","Pie Chart").replace("table","Likert Table").replace("text","Text View")}</div>
                          <button onClick={function() { setShowChartPreview(function(p) { return Object.assign({}, p, {[k]: false}); }); }}
                            style={{ padding: "1px 6px", border: "1px solid #ccc", borderRadius: 3, background: "white", fontSize: 10, cursor: "pointer", color: "#666" }}>✕</button>
                        </div>
                        {renderMiniChart(q, k)}
                      </div>
                    )}
                  </div>;
                })}
              </div>}
            </div>;
          })}
          <button onClick={function() { setView("viz"); }} style={{ marginTop: 4, padding: "9px 24px", background: "#003366", color: "white", border: "none", borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Generate Charts →</button>
        </div>
      )}


      {/* VISUALIZE */}
      {view === "viz" && parsed && data && (
        <div>
          {Object.keys(parsed).map(function(sk) {
            var sec = parsed[sk];
            var activeQs = Object.keys(sec.qs).filter(function(qk) { return picks[sk + "|" + qk]; });
            if (activeQs.length === 0) return null;
            return <div key={sk} style={box}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div>
                  <div style={{ fontSize: 9, letterSpacing: 2, textTransform: "uppercase", color: "#94a3b8" }}>Section {sk}</div>
                  <div style={{ fontSize: 16, fontWeight: 600, color: "#003366" }}>{sec.name}</div>
                </div>
                <button onClick={function() { doNarrative(sk, sec.name); }} disabled={narrLoading[sk]}
                  style={{ padding: "6px 16px", background: narrLoading[sk] ? "#94a3b8" : "#2A9D8F", color: "white", border: "none", borderRadius: 5, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                  {narrLoading[sk] ? "⏳ Generating..." : "✨ Generate Narrative"}
                </button>
              </div>
              {narr[sk] && <div style={{ marginBottom: 16 }}>
                <div style={{ background: "#f0fdfa", padding: "12px 16px", borderRadius: "6px 6px 0 0", borderLeft: "4px solid #2A9D8F", fontSize: 13, lineHeight: 1.6, color: "#1e293b" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                    <div style={{ fontSize: 9, letterSpacing: 2, textTransform: "uppercase", color: "#2A9D8F", fontWeight: 600 }}>AI Narrative</div>
                    <button onClick={function() { setShowEdit(function(p) { return Object.assign({}, p, {[sk]: !p[sk]}); }); }}
                      style={{ padding: "2px 8px", borderRadius: 4, border: "1px solid #2A9D8F", background: showEdit[sk] ? "#d1fae5" : "transparent", fontSize: 10, cursor: "pointer", color: "#2A9D8F", fontWeight: 600 }}>
                      {showEdit[sk] ? "Hide Editor" : "✏️ Edit / Refine"}
                    </button>
                  </div>
                  <div style={{ lineHeight: 1.7 }}>{renderNarrativeWithRefs(narr[sk])}</div>
                  <div style={{ marginTop: 6, fontSize: 10, color: "#94a3b8" }}>
                    <span style={{ background: "#dbeafe", color: "#1e40af", padding: "0px 4px", borderRadius: 2, fontSize: 9, fontWeight: 600, borderBottom: "2px solid #93c5fd" }}>Qx</span> = data reference
                  </div>
                </div>
                {showEdit[sk] && (
                  <div style={{ background: "#f8fffe", padding: "10px 16px", borderRadius: "0 0 6px 6px", borderLeft: "4px solid #89CFF0", borderTop: "1px dashed #b2dfdb" }}>
                    <div style={{ fontSize: 10, color: "#64748b", marginBottom: 4 }}>Tell the AI how to modify:</div>
                    <div style={{ display: "flex", gap: 8 }}>
                      <textarea value={editPrompt[sk] || ""} onChange={function(e) { var v = e.target.value; setEditPrompt(function(p) { return Object.assign({}, p, {[sk]: v}); }); }}
                        placeholder="e.g. Make shorter, emphasize the 60% finding, translate to Spanish..."
                        style={{ flex: 1, padding: "8px 10px", borderRadius: 5, border: "1px solid #cbd5e1", fontSize: 12, resize: "vertical", minHeight: 50, fontFamily: "system-ui,sans-serif", lineHeight: 1.5 }} />
                      <button onClick={function() { refineNarrative(sk, sec.name); }} disabled={narrLoading[sk] || !(editPrompt[sk] || "").trim()}
                        style={{ padding: "8px 14px", background: narrLoading[sk] ? "#94a3b8" : !(editPrompt[sk] || "").trim() ? "#e2e8f0" : "#003366", color: "white", border: "none", borderRadius: 5, fontSize: 11, fontWeight: 600, cursor: "pointer", alignSelf: "flex-end", whiteSpace: "nowrap" }}>
                        {narrLoading[sk] ? "⏳..." : "🔄 Refine"}
                      </button>
                    </div>
                  </div>
                )}
              </div>}
              {activeQs.map(function(qk) { return renderChart(sec.qs[qk], sk + "|" + qk); })}
            </div>;
          })}
          <div style={{ textAlign: "center", padding: "14px 0 28px", display: "flex", justifyContent: "center", gap: 8, flexWrap: "wrap" }}>
            <button onClick={function() { setView("config"); }} style={{ padding: "8px 20px", background: "white", color: "#003366", border: "2px solid #003366", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>← Reconfigure</button>
            <button onClick={exportWord} style={{ padding: "8px 20px", background: "#2A9D8F", color: "white", border: "none", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>📄 Export Report</button>
            <button onClick={function() { setView("upload"); setData(null); setParsed(null); setNarr({}); setCleanInfo(null); setRawData(null); setDupeRows([]); setDupeChecked({}); setExportHTML(null); setCopyStatus({}); }}
              style={{ padding: "8px 20px", background: "#64748b", color: "white", border: "none", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>New File</button>
          </div>
        </div>
      )}
    </div>
  );
}
