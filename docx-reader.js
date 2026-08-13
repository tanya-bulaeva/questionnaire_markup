function parseNumbering(entries) {
  const entry = entries?.get?.("word/numbering.xml");
  if (!entry) return null;
  const xml = new TextDecoder("utf-8").decode(entry.data);
  const dom = new DOMParser().parseFromString(xml, "application/xml");
  if (dom.querySelector("parsererror")) return null;
  const abstractLevels = new Map();
  Array.from(dom.getElementsByTagNameNS(NS.w, "abstractNum")).forEach((abstractNode) => {
    const abstractId = wAttr(abstractNode, "abstractNumId");
    if (!abstractId) return;
    const levels = new Map();
    Array.from(abstractNode.getElementsByTagNameNS(NS.w, "lvl")).forEach((levelNode) => {
      const level = wAttr(levelNode, "ilvl") || "0";
      const numFmt = wAttr(firstChildByLocalName(levelNode, "numFmt"), "val") || "";
      const start = Number(wAttr(firstChildByLocalName(levelNode, "start"), "val") || "1");
      levels.set(level, {
        numFmt,
        start: Number.isInteger(start) && start > 0 ? start : 1
      });
    });
    abstractLevels.set(abstractId, levels);
  });
  const nums = new Map();
  Array.from(dom.getElementsByTagNameNS(NS.w, "num")).forEach((numNode) => {
    const numId = wAttr(numNode, "numId");
    const abstractId = wAttr(firstChildByLocalName(numNode, "abstractNumId"), "val");
    if (numId && abstractId) nums.set(numId, abstractId);
  });
  return { nums, abstractLevels };
}

function paragraphNumberSpec(paragraph, numbering) {
  const pPr = firstChildByLocalName(paragraph, "pPr");
  const numPr = pPr && firstChildByLocalName(pPr, "numPr");
  if (!numPr) return null;
  const numId = wAttr(firstChildByLocalName(numPr, "numId"), "val");
  const level = wAttr(firstChildByLocalName(numPr, "ilvl"), "val") || "0";
  if (!numId) return null;
  const abstractId = numbering?.nums?.get?.(numId);
  const levelSpec = abstractId ? numbering?.abstractLevels?.get?.(abstractId)?.get?.(level) : null;
  if (levelSpec && levelSpec.numFmt && levelSpec.numFmt !== "decimal") return null;
  return { numId, level, start: levelSpec?.start || 1 };
}

function nextParagraphListNumber(spec, counters) {
  if (!spec) return "";
  const key = `${spec.numId}:${spec.level}`;
  const previous = counters.get(key);
  const next = previous == null ? spec.start : previous + 1;
  counters.set(key, next);
  Array.from(counters.keys()).forEach((counterKey) => {
    const [numId, level] = counterKey.split(":");
    if (numId === spec.numId && Number(level) > Number(spec.level)) counters.delete(counterKey);
  });
  return String(next);
}

function extractBlocks(dom, numbering = null) {
  const body = firstChildByLocalName(dom.documentElement, "body");
  const blocks = [];
  const numberingCounters = new Map();
  if (!body) return blocks;
  const appendBlockFromNode = (node) => {
    if (node.nodeType !== 1) return;
    const local = node.localName;
    if (local === "p") {
      const runs = trimFormattedRuns(getFormattedRuns(node));
      const text = runs.map(run => run.text).join("");
      if (!text) return;
      const listNumber = nextParagraphListNumber(paragraphNumberSpec(node, numbering), numberingCounters);
      const changeMarks = blockChangeMarks(node, runs);
      blocks.push({
        id: blocks.length,
        kind: "p",
        text,
        listNumber,
        runs,
        change_mark: changeMarkFromRuns(runs),
        change_marks: changeMarks.colors,
        strike_ratio: changeMarks.strikeRatio,
        htmlText: renderFormattedRuns(runs, { mode: "block" }),
        table: null,
        xmlIndex: getElementIndex(node),
        role: "unassigned"
      });
      return;
    } else if (local === "tbl") {
      const table = extractTable(node);
      const text = table.rows.map(row => row.join(" | ")).join(" / ").trim();
      if (!text) return;
      const runs = trimFormattedRuns(getFormattedRuns(node));
      const changeMarks = blockChangeMarks(node, runs);
      blocks.push({
        id: blocks.length,
        kind: "tbl",
        text,
        runs,
        change_mark: changeMarkFromRuns(runs),
        change_marks: changeMarks.colors,
        strike_ratio: changeMarks.strikeRatio,
        htmlText: escapeHtml(text),
        table,
        xmlIndex: getElementIndex(node),
        role: "unassigned"
      });
      return;
    }
    Array.from(node.childNodes || []).forEach(appendBlockFromNode);
  };
  Array.from(body.childNodes).forEach(appendBlockFromNode);
  return blocks;
}

function extractTable(tbl) {
  const rows = [];
  Array.from(tbl.childNodes).forEach((tr) => {
    if (tr.nodeType !== 1 || tr.localName !== "tr") return;
    const row = [];
    Array.from(tr.childNodes).forEach((tc) => {
      if (tc.nodeType !== 1 || tc.localName !== "tc") return;
      row.push(getText(tc).replace(/\s+/g, " ").trim());
    });
    if (row.some(Boolean)) rows.push(row);
  });
  return { rows };
}

function uniqueStrings(values) {
  return [...new Set((values || []).map(value => String(value || "").trim()).filter(Boolean))];
}

function firstChildByLocalName(node, localName) {
  return Array.from(node.childNodes).find(child => child.nodeType === 1 && child.localName === localName) || null;
}

function wAttr(node, name) {
  if (!node) return "";
  return node.getAttributeNS(NS.w, name) || node.getAttribute(`w:${name}`) || node.getAttribute(name) || "";
}

function getElementIndex(node) {
  return Array.from(node.parentNode.childNodes).filter(n => n.nodeType === 1).indexOf(node);
}

function getText(node) {
  const parts = [];
  const walk = (current) => {
    Array.from(current.childNodes).forEach((child) => {
      if (child.nodeType !== 1) return;
      if (child.namespaceURI === NS.w) {
        if (child.localName === "t") {
          parts.push(child.textContent || "");
          return;
        }
        if (child.localName === "tab") {
          parts.push("\t");
          return;
        }
        if (child.localName === "br" || child.localName === "cr") {
          parts.push(" ");
          return;
        }
        if (child.localName === "noBreakHyphen" || child.localName === "softHyphen") {
          parts.push("-");
          return;
        }
      }
      walk(child);
    });
  };
  walk(node);
  return parts.join("");
}

function getFormattedRuns(node) {
  const runs = [];
  const walk = (current) => {
    Array.from(current.childNodes).forEach((child) => {
      if (child.nodeType !== 1) return;
      if (child.namespaceURI === NS.w && child.localName === "r") {
        const text = getText(child);
        if (text) runs.push({
          text,
          bold: isRunPropertyOn(child, "b"),
          italic: isRunPropertyOn(child, "i"),
          underline: hasRunProperty(child, "u"),
          color: runColor(child),
          highlight: runHighlight(child),
          fill: runFill(child),
          strike: isRunPropertyOn(child, "strike") || isRunPropertyOn(child, "dstrike")
        });
        return;
      }
      walk(child);
    });
  };
  walk(node);
  return mergeFormattedRuns(runs);
}

function hasRunProperty(run, name) {
  const rPr = firstChildByLocalName(run, "rPr");
  return !!(rPr && firstChildByLocalName(rPr, name));
}

function isRunPropertyOn(run, name) {
  const rPr = firstChildByLocalName(run, "rPr");
  const property = rPr && firstChildByLocalName(rPr, name);
  if (!property) return false;
  const value = property.getAttributeNS(NS.w, "val");
  return value !== "false" && value !== "0";
}

function runColor(run) {
  const rPr = firstChildByLocalName(run, "rPr");
  const color = rPr && firstChildByLocalName(rPr, "color");
  const value = color?.getAttributeNS(NS.w, "val") || "";
  return normalizeColor(value);
}

function runHighlight(run) {
  const rPr = firstChildByLocalName(run, "rPr");
  const highlight = rPr && firstChildByLocalName(rPr, "highlight");
  return String(highlight?.getAttributeNS(NS.w, "val") || "").trim().toLowerCase();
}

function runFill(run) {
  const rPr = firstChildByLocalName(run, "rPr");
  const shd = rPr && firstChildByLocalName(rPr, "shd");
  const value = shd?.getAttributeNS(NS.w, "fill") || "";
  return normalizeColor(value);
}

function paragraphFill(paragraph) {
  const pPr = firstChildByLocalName(paragraph, "pPr");
  const shd = pPr && firstChildByLocalName(pPr, "shd");
  return normalizeColor(shd?.getAttributeNS(NS.w, "fill") || "");
}

function blockChangeMarks(node, runs) {
  const colors = [];
  blockFillColors(node).forEach(color => colors.push(color));
  (runs || []).forEach((run) => {
    const highlight = String(run.highlight || "").toLowerCase();
    if (highlight && highlight !== "none" && highlight !== "white") colors.push(highlight);
    if (isVisibleChangeFill(run.fill)) colors.push(run.fill);
  });
  const total = (runs || []).reduce((sum, run) => sum + visibleCharCount(run.text), 0);
  const struck = (runs || []).reduce((sum, run) => sum + (run.strike ? visibleCharCount(run.text) : 0), 0);
  return {
    colors: uniqueStrings(colors),
    strikeRatio: total ? struck / total : 0
  };
}

function blockFillColors(node) {
  const colors = [];
  const direct = paragraphFill(node);
  if (isVisibleChangeFill(direct)) colors.push(direct);
  Array.from(node?.getElementsByTagNameNS?.(NS.w, "shd") || []).forEach((shd) => {
    const fill = normalizeColor(shd.getAttributeNS(NS.w, "fill") || "");
    if (isVisibleChangeFill(fill)) colors.push(fill);
  });
  return uniqueStrings(colors);
}

function isVisibleChangeFill(color) {
  const value = String(color || "").toUpperCase();
  return !!value && value !== "FFFFFF" && value !== "000000" && value !== "AUTO";
}

function changeMarkFromRuns(runs) {
  let hasMarked = false;
  for (const run of runs || []) {
    const mark = changeMarkFromRun(run);
    if (mark === "red") return "red";
    if (mark === "marked") hasMarked = true;
  }
  return hasMarked ? "marked" : "none";
}

function changeMarkFromRun(run) {
  const highlight = String(run?.highlight || "").toLowerCase();
  if (highlight === "red") return "red";
  if (highlight && highlight !== "none" && highlight !== "white") return "marked";
  const colors = [run?.fill, run?.color].map(value => String(value || "").toUpperCase()).filter(Boolean);
  if (colors.some(isRedChangeColor)) return "red";
  if (colors.some(isMarkedChangeColor)) return "marked";
  return "none";
}

function isRedChangeColor(color) {
  const rgb = color.match(/[0-9A-F]{2}/g);
  if (!rgb) return false;
  const [r, g, b] = rgb.map(part => parseInt(part, 16));
  return r >= 150 && g <= 90 && b <= 90;
}

function isMarkedChangeColor(color) {
  const rgb = color.match(/[0-9A-F]{2}/g);
  if (!rgb) return false;
  const [r, g, b] = rgb.map(part => parseInt(part, 16));
  if (color === "000000" || color === "FFFFFF") return false;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  return max >= 90 && max - min >= 25;
}

function normalizeColor(value) {
  const clean = String(value || "").replace(/^#/, "").trim().toUpperCase();
  return /^[0-9A-F]{6}$/.test(clean) ? clean : "";
}

function colorKind(color) {
  if (!color || color === "000000" || color === "FFFFFF") return { kind: "none", value: "" };
  const rgb = color.match(/[0-9A-F]{2}/g).map(part => parseInt(part, 16));
  const [r, g, b] = rgb;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  if (max < 80) return { kind: "none", value: "" };
  if (max - min < 35) return { kind: "none", value: "" };
  const hue = rgbHue(r, g, b);
  if (hue >= 185 && hue <= 255) return { kind: "blue", value: color };
  return { kind: "color", value: color };
}

function rgbHue(r, g, b) {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const delta = max - min;
  if (!delta) return 0;
  let hue = 0;
  if (max === rn) hue = ((gn - bn) / delta) % 6;
  else if (max === gn) hue = ((bn - rn) / delta) + 2;
  else hue = ((rn - gn) / delta) + 4;
  return (hue * 60 + 360) % 360;
}

function renderFormattedRuns(runs, options = {}) {
  const mode = options.mode || "block";
  const sourceRuns = (Array.isArray(runs) ? runs : []).map(run => ({
    ...run,
    colorInfo: colorKind(run.color)
  }));
  const totalChars = sourceRuns.reduce((sum, run) => sum + visibleCharCount(run.text), 0);
  const boldChars = sourceRuns.reduce((sum, run) => sum + (run.bold ? visibleCharCount(run.text) : 0), 0);
  const plainTitleChars = sourceRuns.reduce((sum, run) => sum + (run.colorInfo.kind === "blue" ? 0 : visibleCharCount(run.text)), 0);
  const boldPlainTitleChars = sourceRuns.reduce((sum, run) => sum + (run.colorInfo.kind === "blue" || !run.bold ? 0 : visibleCharCount(run.text)), 0);
  const suppressBold = totalChars > 0 && (
    (mode === "title" && plainTitleChars > 0 && boldPlainTitleChars === plainTitleChars)
    || ((mode === "info" || mode === "block") && boldChars / totalChars >= 0.85)
  );
  const styledRuns = sourceRuns.map(run => {
    const color = run.colorInfo;
    return {
      text: run.text || "",
      bold: !!run.bold && !suppressBold && color.kind !== "blue",
      italic: !!run.italic,
      underline: !!run.underline,
      colorKind: color.kind,
      colorValue: color.value
    };
  });
  return mergeFormattedRuns(styledRuns).map(run => renderFormattedRun(run)).join("");
}

function renderFormattedRun(run) {
  let text = escapeHtml(run.text || "");
  if (!text) return "";
  if (run.colorKind === "blue") {
    if (run.italic) text = `<i>${text}</i>`;
    if (run.underline) text = `<u>${text}</u>`;
    return `<span class="textblue">${text}</span>`;
  }
  if (run.bold) text = `<b>${text}</b>`;
  if (run.underline) text = `<u>${text}</u>`;
  if (run.italic) text = `<i>${text}</i>`;
  if (run.colorKind === "color") text = `<span style="color:#${run.colorValue.toLowerCase()}">${text}</span>`;
  return text;
}

function mergeFormattedRuns(runs) {
  const merged = [];
  (runs || []).forEach(run => {
    const normalized = {
      text: run.text || "",
      bold: !!run.bold,
      italic: !!run.italic,
      underline: !!run.underline,
      color: run.color || "",
      colorKind: run.colorKind || "",
      colorValue: run.colorValue || "",
      highlight: run.highlight || "",
      fill: run.fill || "",
      strike: !!run.strike
    };
    const prev = merged[merged.length - 1];
    if (prev
      && prev.bold === normalized.bold
      && prev.italic === normalized.italic
      && prev.underline === normalized.underline
      && prev.color === normalized.color
      && prev.colorKind === normalized.colorKind
      && prev.colorValue === normalized.colorValue
      && prev.highlight === normalized.highlight
      && prev.fill === normalized.fill
      && prev.strike === normalized.strike) {
      prev.text += normalized.text;
    } else {
      merged.push(normalized);
    }
  });
  return merged;
}

function sliceFormattedRuns(runs, start, end) {
  const result = [];
  let offset = 0;
  (runs || []).forEach(run => {
    const text = run.text || "";
    const next = offset + text.length;
    if (next > start && offset < end) {
      const from = Math.max(0, start - offset);
      const to = Math.min(text.length, end - offset);
      if (to > from) result.push({ ...run, text: text.slice(from, to) });
    }
    offset = next;
  });
  return mergeFormattedRuns(result);
}

function trimFormattedRuns(runs) {
  const text = (runs || []).map(run => run.text || "").join("");
  const range = trimRange(text, 0, text.length);
  return sliceFormattedRuns(runs, range.start, range.end);
}

function trimRange(text, start, end) {
  let from = Math.max(0, start);
  let to = Math.min(String(text || "").length, end);
  while (from < to && /\s/.test(text[from])) from += 1;
  while (to > from && /\s/.test(text[to - 1])) to -= 1;
  return { start: from, end: to };
}

function rebuildSegmentsFromTexts(block, segmentTexts) {
  let offset = 0;
  const blockText = block.text || "";
  const blockRuns = block.runs || [];
  return (segmentTexts || []).map(text => {
    const clean = String(text || "");
    let start = blockText.indexOf(clean, offset);
    if (start < 0) start = offset;
    const end = start + clean.length;
    offset = end;
    const runs = sliceFormattedRuns(blockRuns, start, end);
    return {
      text: clean,
      runs,
      htmlText: renderFormattedRuns(runs, { mode: "block" })
    };
  });
}

function visibleCharCount(text) {
  return String(text || "").replace(/\s+/g, "").length;
}

async function readZip(bytes) {
      const entries = new Map();
  let offset = 0;
  while (offset + 30 < bytes.length) {
    const sig = readU32(bytes, offset);
    if (sig !== 0x04034b50) break;
    const flags = readU16(bytes, offset + 6);
    const method = readU16(bytes, offset + 8);
    const modTime = readU16(bytes, offset + 10);
    const modDate = readU16(bytes, offset + 12);
    const crc = readU32(bytes, offset + 14);
    let compressedSize = readU32(bytes, offset + 18);
    let uncompressedSize = readU32(bytes, offset + 22);
    const nameLength = readU16(bytes, offset + 26);
    const extraLength = readU16(bytes, offset + 28);
    const nameBytes = bytes.slice(offset + 30, offset + 30 + nameLength);
    const name = new TextDecoder("utf-8").decode(nameBytes);
    const dataStart = offset + 30 + nameLength + extraLength;
    if (flags & 0x08) throw new Error(`ZIP entry uses data descriptor and is not supported: ${name}`);
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);
    let data;
    if (method === 0) data = compressed;
    else if (method === 8) data = await inflateRaw(compressed);
    else throw new Error(`Unsupported ZIP compression method ${method} for ${name}`);
    if (!name.endsWith("/")) {
      entries.set(name, { name, data, modTime, modDate, crc, compressedSize, uncompressedSize, method });
    }
    offset = dataStart + compressedSize;
  }
  if (!entries.size) throw new Error("Не удалось прочитать ZIP внутри DOCX");
  return entries;
}

async function inflateRaw(compressed) {
  if (!("DecompressionStream" in window)) {
    throw new Error("Браузер не поддерживает DecompressionStream. Откройте страницу в актуальном Chrome/Edge.");
  }
  const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
