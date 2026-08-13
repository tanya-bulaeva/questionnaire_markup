async function handleGeoFile(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  setStatus("Читаю GEO XLSX...");
  setBusy(true);
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const entries = await readZip(bytes);
    const rows = parseGeoWorkbook(entries);
    if (!rows.length) throw new Error("В GEO XLSX не найдены строки городов.");
    state.geoFile = file;
    state.geoRows = rows;
    saveGeoRowsToStorage(rows, file);
    state.questions.forEach(q => applyGeoDefaults(q));
    render();
    setStatus(`GEO XLSX загружен: ${rows.length} населенных пунктов.`);
  } catch (error) {
    console.error(error);
    setStatus(`Ошибка GEO XLSX: ${error.message}`);
    alert(error.message);
  } finally {
    setBusy(false);
    event.target.value = "";
  }
}

function geoStorageKey() {
  return "questionnaire-review-geo-v1";
}

function loadBundledGeoRows() {
  try {
    const rows = Array.isArray(window.GEO_ROWS) ? window.GEO_ROWS : [];
    state.geoRows = rows
      .map(normalizeStoredGeoRow)
      .filter(row => row.code && row.city);
    if (state.geoRows.length) {
      state.geoFile = {
        name: "geo-data.js",
        size: 0,
        lastModified: 0
      };
      setStatus(`GEO справочник загружен из geo-data.js: ${state.geoRows.length} населенных пунктов.`);
    }
  } catch (error) {
    console.error(error);
  }
}

function loadGeoRowsFromStorage() {
  try {
    const raw = localStorage.getItem(geoStorageKey());
    if (!raw) return;
    const payload = JSON.parse(raw);
    if (!payload || !Array.isArray(payload.rows) || !payload.rows.length) return;
    state.geoRows = payload.rows
      .map(normalizeStoredGeoRow)
      .filter(row => row.code && row.city);
    state.geoFile = {
      name: payload.source_name || "GEO cache",
      size: 0,
      lastModified: Date.parse(payload.saved_at || "") || 0
    };
    if (state.geoRows.length) {
      setStatus(`GEO справочник загружен из хранилища: ${state.geoRows.length} населенных пунктов.`);
    }
  } catch (error) {
    console.error(error);
  }
}

function saveGeoRowsToStorage(rows, file) {
  try {
    localStorage.setItem(geoStorageKey(), JSON.stringify({
      saved_at: new Date().toISOString(),
      source_name: file?.name || "",
      source_size: file?.size || 0,
      source_last_modified: file?.lastModified || 0,
      rows: rows.map(normalizeStoredGeoRow)
    }));
  } catch (error) {
    console.error(error);
    setStatus("GEO XLSX загружен, но не удалось сохранить справочник в браузере.");
  }
}

function normalizeStoredGeoRow(row) {
  return {
    code: Number(row?.code || 0),
    city: normalizeWhitespace(row?.city || ""),
    region: normalizeWhitespace(row?.region || ""),
    district: normalizeWhitespace(row?.district || ""),
    strata: normalizeWhitespace(row?.strata || ""),
    strata_wi: normalizeWhitespace(row?.strata_wi || ""),
    population: Number(row?.population || 0) || 0,
    district_dop: normalizeWhitespace(row?.district_dop || "")
  };
}

function parseGeoWorkbook(entries) {
  const sharedStrings = parseWorkbookSharedStrings(entries);
  const sheetPath = findWorkbookSheetPath(entries, "ГЕО 2025") || "xl/worksheets/sheet2.xml";
  const sheetEntry = entries.get(sheetPath);
  if (!sheetEntry) throw new Error("В XLSX не найден лист ГЕО 2025.");
  const xml = new TextDecoder("utf-8").decode(sheetEntry.data);
  const dom = new DOMParser().parseFromString(xml, "application/xml");
  return Array.from(dom.getElementsByTagNameNS(NS.wb, "row"))
    .map(row => parseGeoSheetRow(row, sharedStrings))
    .filter(Boolean);
}

function parseWorkbookSharedStrings(entries) {
  const entry = entries.get("xl/sharedStrings.xml");
  if (!entry) return [];
  const xml = new TextDecoder("utf-8").decode(entry.data);
  const dom = new DOMParser().parseFromString(xml, "application/xml");
  return Array.from(dom.getElementsByTagNameNS(NS.wb, "si")).map((si) =>
    Array.from(si.getElementsByTagNameNS(NS.wb, "t")).map(t => t.textContent || "").join("")
  );
}

function findWorkbookSheetPath(entries, sheetName) {
  const workbookEntry = entries.get("xl/workbook.xml");
  const relsEntry = entries.get("xl/_rels/workbook.xml.rels");
  if (!workbookEntry || !relsEntry) return "";
  const workbook = new DOMParser().parseFromString(new TextDecoder("utf-8").decode(workbookEntry.data), "application/xml");
  const rels = new DOMParser().parseFromString(new TextDecoder("utf-8").decode(relsEntry.data), "application/xml");
  const relById = new Map(Array.from(rels.documentElement.children).map(rel => [rel.getAttribute("Id"), rel.getAttribute("Target")]));
  const sheet = Array.from(workbook.getElementsByTagNameNS(NS.wb, "sheet")).find(item => item.getAttribute("name") === sheetName);
  const relId = sheet?.getAttributeNS(NS.relOffice, "id") || sheet?.getAttribute("r:id");
  const target = relById.get(relId || "");
  if (!target) return "";
  return target.startsWith("xl/") ? target : `xl/${target}`;
}

function parseGeoSheetRow(row, sharedStrings) {
  const rowIndex = Number(row.getAttribute("r") || 0);
  if (rowIndex < 4) return null;
  const cells = new Map(Array.from(row.getElementsByTagNameNS(NS.wb, "c")).map(cell => [xlsxColumnIndex(cell.getAttribute("r") || ""), xlsxCellValue(cell, sharedStrings)]));
  const code = Number(cells.get(1));
  const city = normalizeWhitespace(cells.get(2));
  if (!Number.isFinite(code) || !city) return null;
  return {
    code,
    city,
    region: normalizeWhitespace(cells.get(3)),
    district: normalizeWhitespace(cells.get(4)),
    strata: normalizeWhitespace(cells.get(5)),
    strata_wi: normalizeWhitespace(cells.get(6)),
    population: Number(cells.get(7) || 0) || 0,
    district_dop: normalizeWhitespace(cells.get(8))
  };
}

function xlsxColumnIndex(cellRef) {
  const match = String(cellRef || "").match(/^([A-Z]+)/i);
  if (!match) return 0;
  return match[1].toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}

function xlsxCellValue(cell, sharedStrings) {
  const value = firstChildByLocalName(cell, "v")?.textContent || "";
  if (cell.getAttribute("t") === "s") return sharedStrings[Number(value)] || "";
  return value;
}

const GEO_CITY_HTML1_LINES = [
  "<style> .es__question-badge { line-height: normal; } .px-3:first-child .question-error-message.question-message { display: none; } .px-3.py-2.mb-2 { padding-bottom: 0!important; margin-bottom: 0!important; } .es__answer-item { margin-bottom: 0; margin-top: 0; } .single-list-container { margin-top: 20px; } .px-3 { padding: 0 !important; } </style> <script> $(function(){ $(\".es__question\").eq(1).hide(); $(\".component-container.container\").css({'padding':'0'}); $('.single-list-container').on('mousedown touchstart', '.select-dropdown-item', (e) => { if (e.originalEvent.detail !== true) { $('.es__answer-item .es__answer-label.active .es__answer-text').each(function (i, node) { node.dispatchEvent(new CustomEvent('click', { detail: true })); }); } }); $(window).on('keydown', (e) => { if (e.key === 'Enter' && $('.listContainer').length) { if (e.originalEvent.detail !== true) { $('.es__answer-item .es__answer-label.active .es__answer-text').each(function (i, node) { node.dispatchEvent(new CustomEvent('click', { detail: true })); }); } } }); jQuery('.es__answer-item .es__answer-label .es__answer-text').on('click', function (e) { $(\".select-clear-button\").trigger('click'); }); }); <\/script>"
];

function geoCityOptions(q) {
  if (String(q?.geo_city_threshold || "") === "custom") {
    return customGeoCityOptions(q);
  }
  const template = geoTemplateForQuestion(q);
  if (template?.cityOptions?.length) {
    return template.cityOptions
      .map(option => ({ code: Number(option.code || 0), city: option.city || "" }))
      .filter(option => option.code !== 9998 && option.code !== 9999);
  }
  if (!state.geoRows.length) return [];
  const threshold = String(q?.geo_city_threshold || "100+");
  const minPopulation = geoThresholdPopulation(threshold);
  return state.geoRows
    .filter(row => threshold === "all" || Number(row.population || 0) >= minPopulation || isGeoAlwaysIncludedRow(row))
    .filter(row => q?.geo_city_include_crimea || !isCrimeaGeoRow(row))
    .filter(row => Number(row.code || 0) !== 9998 && Number(row.code || 0) !== 9999);
}

function customGeoCityOptions(q) {
  return customGeoOptionRows(q)
    .filter(option => Number.isFinite(option.code) && option.city && !isGeoOtherText(option.city));
}

function customGeoOptionRows(q) {
  const rows = [];
  const seenCodes = new Set();
  (Array.isArray(q?.options) ? q.options : []).forEach(option => {
    expandCustomGeoOption(option).forEach(row => {
      const codeKey = String(row.code);
      if (!Number.isFinite(row.code) || !row.city || seenCodes.has(codeKey)) return;
      seenCodes.add(codeKey);
      rows.push(row);
    });
  });
  return rows;
}

function expandCustomGeoOption(option) {
  const code = Number(normalizeWhitespace(option?.code || ""));
  const text = normalizeWhitespace(stripHtmlTags(option?.text || option?.html_text || ""));
  if (!text) return [];
  const slashParts = text.split(/\s*\/\s*/u).map(part => normalizeWhitespace(part)).filter(Boolean);
  const rows = [];
  const addRow = (rowCode, city) => {
    const numericCode = Number(normalizeWhitespace(rowCode || ""));
    const cleanCity = normalizeWhitespace(city || "").replace(/\s*\|\s*$/u, "");
    if (Number.isFinite(numericCode) && cleanCity) rows.push({ code: numericCode, city: cleanCity });
  };
  if (slashParts.length >= 2 && slashParts.some(part => /^\d{1,4}\s*\|/u.test(part))) {
    if (Number.isFinite(code) && !/^\d{1,4}\s*\|/u.test(slashParts[0])) addRow(code, slashParts[0]);
    slashParts.forEach(part => {
      const match = part.match(/^(\d{1,4})\s*\|\s*(.+)$/u);
      if (match) addRow(match[1], match[2]);
    });
    if (rows.length >= 2) return rows;
  }
  return Number.isFinite(code) ? [{ code, city: text }] : [];
}

function geoOtherOption(q) {
  const context = [
    q?.title || "",
    q?.title_html || "",
    ...(q?.instructions || []),
    ...(q?.routing || []).map(item => item?.text || ""),
    ...(q?.source_blocks || []).map(ref => `${refText(ref)} ${refHtml(ref, { mode: "info" })}`)
  ].join(" ");
  const explicit = String(context || "").match(/(?:\+|\bкод\s*)\s*(9998|9999)\s*[.)]?\s*(?:["«“„']?\s*)?(Другой\s+(?:город|населенн?ый\s+пункт))/iu);
  if (explicit) {
    return {
      code: String(explicit[1]),
      text: normalizeGeoOtherText(explicit[2])
    };
  }
  const existing = (q?.options || [])
    .map(option => ({
      code: normalizeWhitespace(option?.code || ""),
      text: normalizeGeoOtherText(option?.text || stripHtmlTags(option?.html_text || ""))
    }))
    .find(option => option.code && isGeoOtherText(option.text));
  if (existing) return existing;
  const badgeText = String(context || "").match(/Другой\s+(?:город|населенн?ый\s+пункт)/iu)?.[0] || "";
  if (badgeText) {
    return {
      code: /(?:\+|\bкод\s*)\s*9998\b/i.test(context) ? "9998" : "9999",
      text: normalizeGeoOtherText(badgeText)
    };
  }
  return { code: "9999", text: "Другой населенный пункт" };
}

function isGeoOtherText(text) {
  const clean = normalizeWhitespace(stripHtmlTags(text || ""))
    .replace(/\s*\([^)]*\)\s*$/u, "")
    .replace(/[.。]+$/u, "");
  return /^(?:Другой\s+(?:город|населенн?ый\s+пункт|регион)|Другое)$/iu.test(clean);
}

function normalizeGeoOtherText(text) {
  const clean = normalizeWhitespace(stripHtmlTags(text || ""))
    .replace(/\s*\([^)]*\)\s*$/u, "")
    .replace(/[.。]+$/u, "");
  if (/^другой\s+город$/iu.test(clean)) return "Другой город";
  if (/^другой\s+населенн?ый\s+пункт$/iu.test(clean)) return "Другой населенный пункт";
  if (/^другой\s+регион$/iu.test(clean)) return "Другой регион";
  if (/^другое$/iu.test(clean)) return "Другое";
  return clean || "Другой населенный пункт";
}

function geoTemplateForQuestion(q) {
  const threshold = String(q?.geo_city_threshold || "");
  if (!threshold || threshold === "custom") return null;
  const key = `${threshold === "all" ? "0+" : threshold}|${q?.geo_city_include_crimea ? "crimea" : "noCrimea"}`;
  return window.GEO_TEMPLATES?.[key] || null;
}

function geoTemplateBlock(q, code) {
  const blocks = geoTemplateForQuestion(q)?.blocks;
  const requestedCode = String(code || "");
  let templateCode = requestedCode;
  let block = blocks?.[templateCode];
  if (!block && blocks && requestedCode) {
    templateCode = Object.keys(blocks).find(key => key.toLowerCase() === requestedCode.toLowerCase()) || requestedCode;
    block = blocks[templateCode];
  }
  if (!block) return "";
  if (requestedCode && templateCode !== requestedCode) {
    block = block.replace(new RegExp(`^${escapeRegExp(templateCode)}(?=\\.)`), requestedCode);
  }
  const actualCityCode = scriptQuestionCode(q);
  return block
    .replace(/\$КОДВОПРОСАГЕО(?![\p{L}\p{N}_])/gu, `$${actualCityCode}`)
    .replace(/КОДВОПРОСАГЕО(?=\.)/g, requestedCode || templateCode || actualCityCode)
    .replace(/\$Qcity\b/g, `$${actualCityCode}`);
}

function geoTemplateCommentLine(q) {
  const template = geoTemplateForQuestion(q);
  if (state.expressMode) return "";
  if (String(q?.geo_city_threshold || "") === "custom") return "//без шаблона";
  if (!template) return "";
  return `//шаблон ${template.threshold} ${template.includeCrimea ? "с Крымом" : "без Крыма"}`;
}

function geoThresholdLabel(threshold) {
  const value = String(threshold || "");
  if (value === "all") return "0+ / все";
  if (value === "custom") return "без шаблона";
  return value || "100+";
}

function geoThresholdPopulation(threshold) {
  if (threshold === "mln+") return 1000000;
  if (threshold === "700+") return 700000;
  if (threshold === "500+") return 500000;
  if (threshold === "400+") return 400000;
  if (threshold === "250+") return 250000;
  if (threshold === "50+") return 50000;
  if (threshold === "all") return 0;
  return 100000;
}

function isCrimeaGeoRow(row) {
  return /крым|севастополь/iu.test(`${row?.region || ""} ${row?.city || ""}`);
}

function isGeoAlwaysIncludedRow(row) {
  return row?.strata === "Мск" || row?.strata === "Спб";
}

const GEO_REGION_CODE_BY_NAME = {
  "Адыгея": 1, "Алтайский край": 2, "Амурская область": 3, "Архангельская область": 4, "Астраханская область": 5,
  "Башкортостан": 6, "Белгородская область": 7, "Брянская область": 8, "Бурятия": 9, "Владимирская область": 10,
  "Волгоградская область": 11, "Вологодская область": 12, "Воронежская область": 13, "Дагестан": 14, "Еврейская АО": 15,
  "Забайкальский край": 16, "Ивановская область": 17, "Ингушетия": 18, "Иркутская область": 19, "Кабардино-Балкария": 20,
  "Калининградская область": 21, "Калмыкия": 22, "Калужская область": 23, "Камчатский край": 24, "Карачаево-Черкесия": 25,
  "Карелия": 26, "Кемеровская область": 27, "Кировская область": 28, "Коми": 29, "Костромская область": 30,
  "Краснодарский край": 31, "Красноярский край": 32, "Курганская область": 33, "Курская область": 34, "Ленинградская область": 35,
  "Липецкая область": 36, "Магаданская область": 37, "Марий Эл": 38, "Мордовия": 39, "Большая Москва": 40,
  "Московская область": 41, "Мурманская область": 42, "Ненецкий Авт. Окр.": 43, "Нижегородская область": 44, "Новгородская область": 45,
  "Новосибирская область": 46, "Омская область": 47, "Оренбургская область": 48, "Орловская область": 49, "Пензенская область": 50,
  "Пермский край": 51, "Приморский край": 52, "Псковская область": 53, "Республика Алтай": 54, "Республика Крым": 55,
  "Ростовская область": 56, "Рязанская область": 57, "Самарская область": 58, "Санкт-Петербург": 59, "Саратовская область": 60,
  "Сахалинская область": 61, "Свердловская область": 62, "Севастополь": 63, "Северная Осетия - Алания": 64, "Смоленская область": 65,
  "Ставропольский край": 66, "Тамбовская область": 67, "Татарстан": 68, "Тверская область": 69, "Томская область": 70,
  "Тульская область": 71, "Тыва": 72, "Тюменская область": 73, "Удмуртия": 74, "Ульяновская область": 75,
  "Хабаровский край": 76, "Хакасия": 77, "Ханты-Мансийский АО": 78, "Челябинская область": 79, "Чечня": 80,
  "Ханты-Мансийский авт. окр.": 78, "Чувашия": 81, "Чукотский АО": 82, "Чукотский авт. окр.": 82, "Якутия": 83,
  "Ямало-Ненецкий АО": 84, "Ямало-Ненецкий авт. окр.": 84, "Ярославская область": 85
};

const GEO_DISTRICT_CODE_BY_NAME = {
  "ЦФО": 1,
  "СЗФО": 2,
  "ЮФО": 3,
  "СКФО": 4,
  "ПФО": 5,
  "УФО": 6,
  "СФО": 7,
  "ДФО": 8,
  "ДВФО": 8
};

const GEO_DISTRICT_DOP_CODE_BY_NAME = {
  "Мск": 1,
  "Спб": 2,
  "ЦФО без Мск": 3,
  "СЗФО без Спб": 4,
  "ЮФО": 5,
  "СКФО": 6,
  "ПФО": 7,
  "УФО": 8,
  "СФО": 9,
  "ДВФО": 10
};

const GEO_STRATA_WI_CODE_BY_NAME = {
  "Мск": 1,
  "Спб": 2,
  "700+": 3,
  "400-700": 4,
  "100-400": 5,
  "100-": 6
};

function geoAutocodeTargets(q) {
  const allowed = ["region", "district", "district_dop", "strata", "strata_wi"];
  if (q?.geo_autocode_targets_manual) {
    return new Set((Array.isArray(q?.geo_autocode_targets) ? q.geo_autocode_targets : [])
      .filter(target => allowed.includes(target)));
  }
  const explicitTargets = explicitGeoAutocodeTargets(q);
  const storedTargets = new Set((Array.isArray(q?.geo_autocode_targets) ? q.geo_autocode_targets : [])
    .filter(target => allowed.includes(target)));
  return new Set([...explicitTargets, ...storedTargets]);
}

function geoAutocodeUiTargets(q) {
  return geoAutocodeTargets(q);
}

function explicitGeoAutocodeTargets(q) {
  const text = normalizeWhitespace([
    q?.title || "",
    ...(q?.instructions || []),
    ...(q?.routing || []).map(item => item?.text || ""),
    ...(q?.source_blocks || []).map(ref => refText(ref))
  ].join(" "));
  const targets = new Set();
  if (/\bQregion(?![A-Za-z0-9_])|закодир\w*\s+регион|субъект\s+рф/iu.test(text)) targets.add("region");
  if (/\bQdistrictDop(?![A-Za-z0-9_])|фо[\s\S]{0,80}(?:мск|москв)[\s\S]{0,80}спб|(?:мск|москв)[\s\S]{0,80}спб[\s\S]{0,80}фо/iu.test(text)) targets.add("district_dop");
  if (/\bQdistrict(?![A-Za-z0-9_])|закодир\w*\s+(?:фо|федеральн\w+\s+округ)|федеральн\w+\s+округ|(?:цфо|сзфо|юфо|скфо|пфо|уфо|сфо|дфо|двфо)/iu.test(text)) targets.add("district");
  if (/\bQstrataWI(?![A-Za-z0-9_])|webindex|web\s*index|кодирование\s+по\s+webindex/iu.test(text)) targets.add("strata_wi");
  if (/\bQstrata(?![A-Za-z0-9_])|закодир\w*\s+страт|страты?|размер\s+города|по\s+интервалам/iu.test(text)) targets.add("strata");
  if (targets.has("district_dop") && !/\bQdistrict(?![A-Za-z0-9_])/i.test(text)) targets.delete("district");
  return targets;
}

function explicitGeoAutocodeLabel(q, target, fallback) {
  const text = [
    q?.title || "",
    ...(q?.instructions || []),
    ...(q?.routing || []).map(item => item?.text || ""),
    ...(q?.source_blocks || []).map(ref => refText(ref))
  ].join(" ");
  const patternByTarget = {
    region: /\bQregion(?![A-Za-z0-9_])/gi,
    district: /\bQdistrict(?![A-Za-z0-9_])/gi,
    district_dop: /\bQdistrictDop(?![A-Za-z0-9_])/gi,
    strata: /\bQstrata(?![A-Za-z0-9_])/gi,
    strata_wi: /\bQstrataWI(?![A-Za-z0-9_])/gi
  };
  const match = String(text || "").match(patternByTarget[target]);
  return match?.[0] || fallback;
}

function buildGeoRegionQuestionBlock(q) {
  const code = explicitGeoAutocodeLabel(q, "region", "Qregion");
  const templateBlock = geoTemplateBlock(q, code);
  if (templateBlock) return templateBlock;
  const options = geoCityOptions(q);
  if (!options.length) return "";
  const groups = groupGeoRows(options, row => geoRegionCode(row), row => row.region);
  if (!groups.length) return "";
  return [
    `${code}. В каком регионе вы проживаете в настоящее время?`,
    "type: Single",
    "template: list",
    "search: yes",
    "before: <code>",
    "$city=[",
    ...geoPhpArrayLines(groups),
    "];",
    `$answer=$q->util()->searchInArrays($${scriptQuestionCode(q)}, $city);`,
    'if ($answer) $q->answer($answer)->next();',
    'else $q->finish("screenout", null, "NO_".$q->questionCode);',
    "</code>",
    "options:",
    ...groups.map(group => `${group.code}. ${group.label}`)
  ].join("\n");
}

function buildGeoDistrictQuestionBlock(q) {
  const code = explicitGeoAutocodeLabel(q, "district", "Qdistrict");
  const templateBlock = geoTemplateBlock(q, code);
  if (templateBlock) return templateBlock;
  const options = geoCityOptions(q);
  if (!options.length) return "";
  const primaryGroups = groupGeoRows(options, row => geoDistrictCode(row), row => row.district);
  const groups = primaryGroups.length ? primaryGroups : groupGeoRows(options, row => geoDistrictDopCode(row), row => row.district_dop);
  if (!groups.length) return "";
  return [
    `${code}. Федеральный округ`,
    "type: Single",
    "before: <code>",
    "$city=[",
    ...geoPhpArrayLines(groups),
    "];",
    `$answer=$q->util()->searchInArrays($${scriptQuestionCode(q)}, $city);`,
    'if ($answer) $q->answer($answer)->next();',
    'else $q->finish("screenout", null, "NO_".$q->questionCode);',
    "</code>",
    "options:",
    ...groups.map(group => `${group.code}. ${group.label}`)
  ].join("\n");
}

function buildGeoDistrictDopQuestionBlock(q) {
  const code = explicitGeoAutocodeLabel(q, "district_dop", "QdistrictDop");
  const templateBlock = geoTemplateBlock(q, code);
  if (templateBlock) return templateBlock;
  const options = geoCityOptions(q);
  if (!options.length) return "";
  const groups = groupGeoRows(options, row => geoDistrictDopCode(row), row => row.district_dop);
  if (!groups.length) return "";
  return [
    `${code}. ФО (МСК и СПБ отдельно)`,
    "type: Single",
    "before: <code>",
    "$city=[",
    ...geoPhpArrayLines(groups),
    "];",
    `$answer=$q->util()->searchInArrays($${scriptQuestionCode(q)}, $city);`,
    'if ($answer) $q->answer($answer)->next();',
    'else $q->finish("screenout", null, "NO_".$q->questionCode);',
    "</code>",
    "options:",
    ...groups.map(group => `${group.code}. ${group.label}`)
  ].join("\n");
}

function buildGeoStrataQuestionBlock(q) {
  const code = explicitGeoAutocodeLabel(q, "strata", "Qstrata");
  const templateBlock = geoTemplateBlock(q, code);
  if (templateBlock) return templateBlock;
  const options = geoCityOptions(q);
  if (!options.length) return "";
  const groups = groupGeoRows(options, row => geoStrataCode(row), row => geoStrataLabel(row));
  if (!groups.length) return "";
  return [
    `${code}. Размер города`,
    "type: Single",
    "before: <code>",
    "$city=[",
    ...geoPhpArrayLines(groups),
    "];",
    `$answer=$q->util()->searchInArrays($${scriptQuestionCode(q)}, $city);`,
    'if ($answer) $q->answer($answer)->next();',
    'else $q->finish("screenout", null, "NO_".$q->questionCode);',
    "</code>",
    "options:",
    ...groups.map(group => `${group.code}. ${group.label}`)
  ].join("\n");
}

function buildGeoStrataWiQuestionBlock(q) {
  const code = explicitGeoAutocodeLabel(q, "strata_wi", "QstrataWI");
  const templateBlock = geoTemplateBlock(q, code);
  if (templateBlock) return templateBlock;
  const options = geoCityOptions(q);
  if (!options.length) return "";
  const groups = groupGeoRows(options, row => geoStrataWiCode(row), row => row.strata_wi);
  if (!groups.length) return "";
  return [
    `${code}. Кодирование по WebIndex`,
    "type: Single",
    "before: <code>",
    "$city=[",
    ...geoPhpArrayLines(groups),
    "];",
    `$answer=$q->util()->searchInArrays($${scriptQuestionCode(q)}, $city);`,
    'if ($answer) $q->answer($answer)->next();',
    'else $q->finish("screenout", null, "NO_".$q->questionCode);',
    "</code>",
    "options:",
    ...groups.map(group => `${group.code}. ${group.label}`)
  ].join("\n");
}

function groupGeoRows(rows, codeFn, labelFn) {
  const byCode = new Map();
  rows.forEach((row) => {
    const code = codeFn(row);
    const label = labelFn(row);
    if (!Number.isFinite(code) || !label) return;
    if (!byCode.has(code)) byCode.set(code, { code, label, cityCodes: [] });
    byCode.get(code).cityCodes.push(Number(row.code));
  });
  return Array.from(byCode.values()).sort((left, right) => left.code - right.code);
}

function geoPhpArrayLines(groups) {
  return groups.map((group, index) => `${group.code} => [${group.cityCodes.join(", ")}]${index + 1 < groups.length ? "," : ""}`);
}

function geoRegionCode(row) {
  return GEO_REGION_CODE_BY_NAME[row?.region || ""] || null;
}

function geoDistrictCode(row) {
  return GEO_DISTRICT_CODE_BY_NAME[row?.district || ""] || null;
}

function geoDistrictDopCode(row) {
  return GEO_DISTRICT_DOP_CODE_BY_NAME[row?.district_dop || ""] || null;
}

function geoStrataCode(row) {
  const label = geoStrataLabel(row);
  const codes = { "Мск": 1, "Спб": 2, "mln+": 3, "500-mln": 4, "250-500": 5, "100-250": 6, "50-100": 7, "<50": 8 };
  return codes[label] || null;
}

function geoStrataWiCode(row) {
  return GEO_STRATA_WI_CODE_BY_NAME[row?.strata_wi || ""] || null;
}

function geoStrataLabel(row) {
  if (row?.strata === "Мск") return "Мск";
  if (row?.strata === "Спб") return "Спб";
  if (row?.strata === "mln+") return "mln+";
  return row?.strata || "";
}

function customGeoScriptOptionLines(q) {
  const rows = customGeoOptionRows(q);
  const lines = rows
    .map(row => ({ ...row, city: cleanCustomGeoOptionText(row.city) }))
    .filter(row => !isGeoOtherText(row.city))
    .map(row => `${row.code}. ${decodeScriptQuoteEntities(escapeHtml(row.city))}`);
  const existingOther = rows
    .map(row => ({ ...row, city: cleanCustomGeoOptionText(row.city) }))
    .find(row => isGeoOtherText(row.city));
  if (existingOther) {
    lines.push(`${existingOther.code}. ${decodeScriptQuoteEntities(escapeHtml(normalizeGeoOtherText(existingOther.city)))}`);
  } else if (hasExplicitCustomGeoOther(q)) {
    lines.push("9999. Другое");
  }
  return lines;
}

function cleanCustomGeoOptionText(text) {
  return normalizeGeoOtherText(stripScreenoutServiceSuffix(text || ""));
}

function hasExplicitCustomGeoOther(q) {
  const text = normalizeWhitespace([
    q?.title || "",
    ...(q?.instructions || []),
    ...(q?.source_blocks || []).map(ref => refText(ref)),
    ...(Array.isArray(q?.options) ? q.options : []).map(option => `${option?.code || ""} ${option?.text || ""}`)
  ].join(" "));
  return /(?:^|\s)\+\s*Другое\b/iu.test(text)
    || /Другой\s+(?:город|населенн?ый\s+пункт|регион)/iu.test(text);
}

function applyGeoDefaults(q) {
  if (!q || (q.kind || "question") === "info") return false;
  if (!isGeoCityCandidate(q)) return false;
  if (q.geo_city_enabled === false) return false;
  const template = inferGeoTemplate(q);
  const wasEnabled = q.geo_city_enabled === true;
  if (!wasEnabled) q.geo_city_enabled = true;
  if (template && !wasEnabled) {
    q.geo_city_threshold = template.threshold;
    q.geo_city_include_crimea = template.includeCrimea;
  } else if (!wasEnabled && hasCustomGeoCityOptions(q)) {
    q.geo_city_threshold = "custom";
    q.geo_city_include_crimea = false;
  } else {
    q.geo_city_threshold = q.geo_city_threshold || template?.threshold || "100+";
    q.geo_city_include_crimea = q.geo_city_include_crimea ?? template?.includeCrimea ?? false;
  }
  q.geo_city_other = q.geo_city_other !== false;
  q.type = "Single";
  q.script_options_randomize = false;
  q.script_rotation_mode = "none";
  return true;
}

function isGeoCityQuestion(q) {
  return !!q?.geo_city_enabled && isGeoCityCandidate(q);
}

function isCustomGeoCityQuestion(q) {
  return isGeoCityQuestion(q) && String(q?.geo_city_threshold || "") === "custom";
}

function isGeoCityCandidate(q) {
  const text = normalizeWhitespace([
    q?.code || "",
    q?.normalized_code || "",
    q?.title || "",
    ...(q?.instructions || []),
    ...(q?.source_blocks || []).map(ref => refText(ref))
  ].join(" "));
  const codeText = normalizeWhitespace([q?.code || "", q?.normalized_code || ""].join(" "));
  const hasGeoInstruction = /шаблон\s+гео|список\s+(?:городов|населенн?ых\s+пунктов)|выпадающ\w+\s+список|раскрывающ\w+\s+список|закодир\w*\s+(?:фо|регион|страт|федеральн\w+\s+округ)|Q(?:region|district|strata)/iu.test(text);
  const hasResidenceCityText = /(?:в\s+каком|какой)\s+(?:городе|населенн?ом\s+пункте)[\s\S]{0,120}(?:прожива|живете|живёте)/iu.test(text)
    || /город(?:е)?\/населенн?ом\s+пункте/iu.test(text)
    || /укажите[\s\S]{0,80}(?:ваш\s+)?(?:город|населенн?ый\s+пункт)[\s\S]{0,120}(?:прожива|живете|живёте|живете\s+постоянно|живёте\s+постоянно)?/iu.test(text)
    || /город[\s\S]{0,80}(?:проживания|постоянно(?:го)?\s+прожив)/iu.test(text);
  return /\bQcity\b/i.test(text)
    || /\bCITY\b/i.test(codeText)
    || hasResidenceCityText
    || (hasGeoInstruction && /город|населенн?ый\s+пункт|регион\s+прожив/iu.test(text))
    || (hasCustomGeoCityOptions(q) && /город|населенн?ый\s+пункт/iu.test(text));
}

function hasCustomGeoCityOptions(q) {
  return customGeoCityOptions(q).length >= 2;
}

function inferGeoThreshold(q) {
  const text = normalizeWhitespace([q?.title || "", ...(q?.instructions || []), ...(q?.source_blocks || []).map(ref => refText(ref))].join(" "));
  const explicitTemplateText = text.match(/(?:шаблон\s+гео|список\s+городов\s+с\s+населением)[\s\S]{0,80}/iu)?.[0] || "";
  if (explicitTemplateText) {
    const explicitThreshold = firstGeoThresholdMention(explicitTemplateText);
    if (explicitThreshold) return explicitThreshold;
  }
  return firstGeoThresholdMention(text);
}

function firstGeoThresholdMention(text) {
  const thresholds = [
    ["mln+", /(?:mln|млн|1\s*(?:млн|миллион))\s*\+?/iu],
    ["700+", /\b700\s*(?:тыс\.?|000)?\s*\+/iu],
    ["500+", /\b500\s*(?:тыс\.?|000)?\s*\+/iu],
    ["400+", /\b400\s*(?:тыс\.?|000)?\s*\+/iu],
    ["250+", /\b250\s*(?:тыс\.?|000)?\s*\+/iu],
    ["100+", /\b100\s*(?:тыс\.?|000)?\s*\+/iu],
    ["50+", /\b50\s*(?:тыс\.?|000)?\s*\+/iu],
    ["all", /\b0\s*\+|все\s+город/iu]
  ];
  let best = null;
  thresholds.forEach(([threshold, pattern], order) => {
    const match = String(text || "").match(pattern);
    if (!match) return;
    const index = match.index ?? 0;
    if (!best || index < best.index || (index === best.index && order < best.order)) {
      best = { threshold, index, order };
    }
  });
  return best?.threshold || "";
}

function inferGeoTemplate(q) {
  const threshold = inferGeoThreshold(q);
  const crimea = inferGeoCrimeaMode(q);
  if (!threshold || crimea == null) return null;
  return { threshold, includeCrimea: crimea };
}

function inferGeoCrimeaMode(q) {
  const text = normalizeWhitespace([q?.title || "", ...(q?.instructions || []), ...(q?.source_blocks || []).map(ref => refText(ref))].join(" "));
  if (/без\s+крым/iu.test(text)) return false;
  if (/(?:с\s+крым|крым\s+включ)/iu.test(text)) return true;
  return null;
}

function inferGeoIncludeCrimea(q) {
  return inferGeoCrimeaMode(q) === true;
}
