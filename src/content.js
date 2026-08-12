(function initInfleetOccurrenceExporter() {
  "use strict";

  if (window.__infleetOccurrenceExporterLoaded) {
    return;
  }
  window.__infleetOccurrenceExporterLoaded = true;

  const MESSAGE = {
    PAGE_INFO: "INFLEET_PAGE_INFO",
    EXPORT: "INFLEET_EXPORT",
    APPLY_DATE: "INFLEET_APPLY_DATE_RANGE",
    CANCEL: "INFLEET_CANCEL_EXPORT",
    PROGRESS: "INFLEET_EXPORT_PROGRESS"
  };

  const OCCURRENCE_LINK_SELECTOR = 'a[href*="/occurrences/"][href*="/details"]';
  const RANGE_PATTERN = /\b\d{2}\/\d{2}\/(?:\d{2}|\d{4})\s*-\s*\d{2}\/\d{2}\/(?:\d{2}|\d{4})\b/;
  const DATE_PATTERN = /\b\d{2}\/\d{2}\/\d{4}\b/g;
  const SEVERITY_WORDS = new Set(["Grave", "Moderada", "Media", "Média", "Leve", "Critica", "Crítica", "Alta", "Baixa"]);

  let activeExport = null;
  let overlayCleanup = null;

  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (!request || !request.type) {
      return false;
    }

    if (request.type === MESSAGE.PAGE_INFO) {
      sendResponse(getPageInfo());
      return false;
    }

    if (request.type === MESSAGE.CANCEL) {
      if (activeExport) {
        activeExport.cancelled = true;
      }
      sendResponse({ ok: true });
      return false;
    }

    if (request.type === MESSAGE.APPLY_DATE) {
      applyDateRange(request.payload)
        .then(sendResponse)
        .catch((error) => sendResponse({ ok: false, message: error.message }));
      return true;
    }

    if (request.type === MESSAGE.EXPORT) {
      exportOccurrences(request.payload)
        .then(sendResponse)
        .catch((error) => {
          updateOverlay(`Erro: ${error.message}`, "error");
          sendProgress({ phase: "error", message: error.message });
          sendResponse({ ok: false, message: error.message });
        });
      return true;
    }

    return false;
  });

  function getPageInfo() {
    const rows = collectRenderedRows();
    return {
      ok: true,
      url: location.href,
      isOccurrencesPage: isOccurrencesPage(),
      rangeText: findDateRangeText(),
      dateFilterText: findDateFilterText(),
      selectedStatus: findSelectedStatus(),
      visibleRows: rows.length,
      scrollable: Boolean(findScrollContainer())
    };
  }

  function isOccurrencesPage() {
    return location.hostname === "app.infleet.com.br" && location.pathname.startsWith("/occurrences");
  }

  async function exportOccurrences(options = {}) {
    if (!isOccurrencesPage()) {
      throw new Error("Abra https://app.infleet.com.br/occurrences antes de exportar.");
    }

    if (!window.InfleetXlsx || typeof window.InfleetXlsx.buildWorkbook !== "function") {
      throw new Error("Gerador Excel nao carregou. Recarregue a pagina e tente novamente.");
    }

    const token = { cancelled: false };
    activeExport = token;
    const scrollAll = options.scrollAll !== false;
    const scrollContainer = findScrollContainer();
    const originalTop = scrollContainer ? getScrollTop(scrollContainer) : 0;
    const rowsByKey = new Map();
    const startedAt = new Date();

    try {
      updateOverlay("Preparando leitura...", "running");
      sendProgress({ phase: "running", message: "Preparando leitura...", rows: 0 });

      const status = findSelectedStatus();
      const requestedRangeText = formatRequestedRange(options.startDate, options.endDate);
      const sourceRangeText = findDateRangeText();
      const rangeText = requestedRangeText || sourceRangeText;
      const metadata = {
        status,
        rangeText,
        requestedRangeText,
        sourceRangeText,
        startDate: options.startDate,
        endDate: options.endDate,
        exportedAt: formatDateTime(startedAt)
      };

      const harvest = () => {
        const rows = collectRenderedRows(metadata);
        for (const row of rows) {
          const key = row.id || `${row.date}|${row.vehicle}|${row.type}|${row.detailUrl}`;
          if (!key) {
            continue;
          }
          const current = rowsByKey.get(key);
          rowsByKey.set(key, current ? mergeRows(current, row) : row);
        }
      };

      if (!scrollAll || !scrollContainer) {
        harvest();
      } else {
        await setScrollTop(scrollContainer, 0);
        await waitForRender(450);
        harvest();

        let y = 0;
        let lastMax = Math.max(0, getScrollHeight(scrollContainer) - getClientHeight(scrollContainer));
        let bestKnownMax = lastMax;
        let displayedPercent = 0;
        const step = Math.max(220, Math.floor(getClientHeight(scrollContainer) * 0.72));
        let stableBottomReads = 0;
        let steps = 0;

        while (!token.cancelled) {
          const max = Math.max(0, getScrollHeight(scrollContainer) - getClientHeight(scrollContainer));
          if (y >= max) {
            await setScrollTop(scrollContainer, max);
            await waitForRender(500);
            harvest();
            const refreshedMax = Math.max(0, getScrollHeight(scrollContainer) - getClientHeight(scrollContainer));
            bestKnownMax = Math.max(bestKnownMax, refreshedMax);
            stableBottomReads = refreshedMax === lastMax ? stableBottomReads + 1 : 0;
            lastMax = refreshedMax;
            if (stableBottomReads >= 2) {
              break;
            }
          }

          y = Math.min(Math.max(0, getScrollHeight(scrollContainer) - getClientHeight(scrollContainer)), y + step);
          await setScrollTop(scrollContainer, y);
          await waitForRender(180);
          harvest();
          steps += 1;

          const currentMax = Math.max(0, getScrollHeight(scrollContainer) - getClientHeight(scrollContainer));
          bestKnownMax = Math.max(bestKnownMax, currentMax, y);
          const rawPercent = Math.round((getScrollTop(scrollContainer) / Math.max(1, bestKnownMax)) * 95);
          const percent = Math.max(displayedPercent, Math.min(95, rawPercent));
          displayedPercent = percent;
          const message = `Lendo lista... ${rowsByKey.size} ocorrencias (${percent}%)`;
          updateOverlay(message, "running");
          sendProgress({ phase: "running", message, rows: rowsByKey.size, percent });

          if (steps > 5000) {
            throw new Error("Limite de rolagem atingido. Exporte um periodo menor.");
          }
        }
      }

      if (token.cancelled) {
        updateOverlay("Exportacao cancelada.", "idle");
        return { ok: false, cancelled: true, rows: rowsByKey.size, message: "Exportacao cancelada." };
      }

      const finishingMessage = `Finalizando Excel... ${rowsByKey.size} ocorrencias (99%)`;
      updateOverlay(finishingMessage, "running");
      sendProgress({ phase: "running", message: finishingMessage, rows: rowsByKey.size, percent: 99 });

      const rows = filterRowsByRequestedRange(
        Array.from(rowsByKey.values()),
        options.startDate,
        options.endDate
      );
      if (rows.length === 0) {
        throw new Error("Nenhuma ocorrencia foi encontrada no periodo escolhido.");
      }

      const columns = [
        { key: "date", label: "Data da ocorrencia", width: 18 },
        { key: "groupDateText", label: "Grupo de data", width: 18 },
        { key: "relativeTime", label: "Tempo exibido", width: 18 },
        { key: "type", label: "Tipo de ocorrencia", width: 28 },
        { key: "severity", label: "Severidade", width: 14 },
        { key: "driver", label: "Motorista", width: 24 },
        { key: "vehicle", label: "Veiculo", width: 14 },
        { key: "status", label: "Aba visualizada", width: 20 },
        { key: "rangeText", label: "Periodo exportado", width: 22 },
        { key: "sourceRangeText", label: "Periodo na tela Infleet", width: 24 },
        { key: "id", label: "ID", width: 40 },
        { key: "detailUrl", label: "Link", width: 68 },
        { key: "exportedAt", label: "Exportado em", width: 22 }
      ];
      const bytes = window.InfleetXlsx.buildWorkbook(rows, columns);
      const filename = makeFilename(options.startDate, options.endDate, rangeText);
      downloadBytes(bytes, filename);

      const doneMessage = `Excel gerado com ${rows.length} ocorrencias.`;
      updateOverlay(doneMessage, "success");
      sendProgress({ phase: "done", message: doneMessage, rows: rows.length });

      return { ok: true, rows: rows.length, filename, message: doneMessage };
    } finally {
      if (scrollContainer && options.restoreScroll !== false) {
        await setScrollTop(scrollContainer, originalTop);
      }
      if (activeExport === token) {
        activeExport = null;
      }
    }
  }

  function collectRenderedRows(metadata = {}) {
    const rows = [];
    const seenNodes = new Set();
    const links = Array.from(document.querySelectorAll(OCCURRENCE_LINK_SELECTOR));

    for (const link of links) {
      const rowElement = link.closest("li") || findGridRow(link);
      if (!rowElement || seenNodes.has(rowElement) || !isElementVisible(rowElement)) {
        continue;
      }
      seenNodes.add(rowElement);
      const parsed = parseOccurrenceRow(rowElement, link, metadata);
      if (parsed) {
        rows.push(parsed);
      }
    }

    return rows;
  }

  function findGridRow(node) {
    let current = node.parentElement;
    while (current && current !== document.body) {
      const style = getComputedStyle(current);
      if (style.display === "grid" && current.children.length >= 4) {
        return current;
      }
      current = current.parentElement;
    }
    return null;
  }

  function parseOccurrenceRow(rowElement, link, metadata) {
    const cells = Array.from(rowElement.children).filter((child) => child.nodeType === Node.ELEMENT_NODE);
    if (cells.length < 4) {
      return null;
    }

    const typeCell = cells[0];
    const driverCell = cells[1];
    const vehicleCell = cells[2];
    const dateCell = cells[3];
    const detailUrl = new URL(link.getAttribute("href"), location.origin).href;
    const id = extractOccurrenceId(detailUrl);
    const dateText = normalizeText(dateCell.innerText);
    const dates = dateText.match(DATE_PATTERN) || [];
    const cellDate = dates.length > 0 ? dates[dates.length - 1] : "";
    const groupDateText = findGroupDateTextForRow(rowElement) || brazilDateToGroupText(cellDate);
    const groupDate = groupDateText ? groupDateTextToBrazilDate(groupDateText, cellDate, metadata) : "";
    const date = groupDate || cellDate;
    const relativeTime = cellDate ? normalizeText(dateText.replace(cellDate, "")) : dateText;

    return {
      date,
      groupDateText,
      relativeTime,
      type: extractOccurrenceType(typeCell),
      severity: extractSeverity(typeCell),
      driver: extractDriver(driverCell),
      vehicle: extractVehicle(vehicleCell),
      status: metadata.status || findSelectedStatus(),
      rangeText: metadata.rangeText || findDateRangeText(),
      sourceRangeText: metadata.sourceRangeText || findDateRangeText(),
      id,
      detailUrl,
      exportedAt: metadata.exportedAt || formatDateTime(new Date())
    };
  }

  function findGroupDateTextForRow(rowElement) {
    const rowWrapper = rowElement.parentElement;
    const rowOffset = readTranslateYOffset(rowWrapper);
    const parent = rowWrapper?.parentElement;

    if (parent && Number.isFinite(rowOffset)) {
      const candidates = Array.from(parent.children)
        .map((child) => ({
          child,
          offset: readTranslateYOffset(child),
          heading: child.querySelector("h2")
        }))
        .filter((candidate) => candidate.heading && Number.isFinite(candidate.offset))
        .map((candidate) => ({
          offset: candidate.offset,
          text: normalizeText(candidate.heading.innerText || candidate.heading.textContent || "")
        }))
        .filter((candidate) => candidate.offset <= rowOffset && parseGroupDateText(candidate.text))
        .sort((a, b) => b.offset - a.offset);

      if (candidates.length > 0) {
        return candidates[0].text;
      }
    }

    const rowRect = rowElement.getBoundingClientRect();
    const visualCandidate = Array.from(document.querySelectorAll("h2"))
      .filter(isElementVisible)
      .map((heading) => ({
        text: normalizeText(heading.innerText || heading.textContent || ""),
        top: heading.getBoundingClientRect().top
      }))
      .filter((candidate) => candidate.top <= rowRect.top && parseGroupDateText(candidate.text))
      .sort((a, b) => b.top - a.top)[0];

    return visualCandidate?.text || "";
  }

  function readTranslateYOffset(element) {
    if (!element) {
      return Number.NaN;
    }

    const transform = element.style.transform || getComputedStyle(element).transform || "";
    let match = transform.match(/translateY\((-?\d+(?:\.\d+)?)px\)/);
    if (match) {
      return Number(match[1]);
    }

    match = transform.match(/matrix\([^,]+,[^,]+,[^,]+,[^,]+,[^,]+,\s*(-?\d+(?:\.\d+)?)\)/);
    return match ? Number(match[1]) : Number.NaN;
  }

  function groupDateTextToBrazilDate(groupDateText, fallbackCellDate, metadata = {}) {
    const parsed = parseGroupDateText(groupDateText);
    if (!parsed) {
      return "";
    }

    const fallbackIso = brazilDateToIso(fallbackCellDate);
    const fallbackParts = fallbackIso ? isoDateParts(fallbackIso) : null;
    let year = fallbackParts && fallbackParts.month === parsed.month ? fallbackParts.year : "";
    if (!year) {
      year = inferYearForGroupDate(parsed.day, parsed.month, metadata);
    }

    return `${String(parsed.day).padStart(2, "0")}/${String(parsed.month).padStart(2, "0")}/${year}`;
  }

  function brazilDateToGroupText(dateText) {
    const isoDate = brazilDateToIso(dateText);
    if (!isoDate) {
      return "";
    }

    const parts = isoDateParts(isoDate);
    return `${parts.day} de ${monthNamePt(parts.month)}`;
  }

  function parseGroupDateText(text) {
    const comparable = normalizeForCompare(text);
    const match = comparable.match(/^(\d{1,2})\s+(?:de\s+)?([a-z]+)$/);
    if (!match) {
      return null;
    }

    const month = monthNumberFromName(match[2]);
    if (!month) {
      return null;
    }

    return {
      day: Number(match[1]),
      month
    };
  }

  function monthNumberFromName(monthName) {
    const months = [
      "janeiro",
      "fevereiro",
      "marco",
      "abril",
      "maio",
      "junho",
      "julho",
      "agosto",
      "setembro",
      "outubro",
      "novembro",
      "dezembro"
    ];
    return months.indexOf(monthName) + 1;
  }

  function inferYearForGroupDate(day, month, metadata = {}) {
    const years = [metadata.startDate, metadata.endDate]
      .filter(isIsoDate)
      .map((date) => isoDateParts(date).year)
      .filter(Boolean);
    const uniqueYears = Array.from(new Set(years));
    const dayText = String(day).padStart(2, "0");
    const monthText = String(month).padStart(2, "0");

    for (const year of uniqueYears) {
      const candidate = `${year}-${monthText}-${dayText}`;
      if ((!metadata.startDate || candidate >= metadata.startDate) && (!metadata.endDate || candidate <= metadata.endDate)) {
        return year;
      }
    }

    return uniqueYears[0] || String(new Date().getFullYear());
  }

  function isoDateParts(isoDate) {
    const [year, month, day] = String(isoDate || "").split("-");
    return {
      year,
      month: Number(month),
      day: Number(day)
    };
  }

  function extractOccurrenceId(url) {
    const match = String(url).match(/\/occurrences\/([^/]+)\/details/);
    return match ? decodeURIComponent(match[1]) : "";
  }

  function extractOccurrenceType(cell) {
    const title = cell.querySelector(".text-text-title");
    if (title) {
      return normalizeText(title.innerText);
    }

    const spans = Array.from(cell.querySelectorAll("span"));
    const candidate = spans
      .map((span) => normalizeText(span.innerText))
      .find((text) => text && !SEVERITY_WORDS.has(text) && text.length > 1);

    if (candidate) {
      return candidate;
    }

    return normalizeText(cell.innerText)
      .replace(/\b(Grave|Moderada|Média|Media|Leve|Crítica|Critica|Alta|Baixa)\b/g, "")
      .trim();
  }

  function extractSeverity(cell) {
    const spans = Array.from(cell.querySelectorAll("span"));
    for (const span of spans) {
      const text = normalizeText(span.innerText);
      if (SEVERITY_WORDS.has(text)) {
        return text;
      }
    }
    return "";
  }

  function extractDriver(cell) {
    const text = normalizeText(cell.innerText);
    return text === "?" ? "" : text;
  }

  function extractVehicle(cell) {
    const texts = Array.from(cell.querySelectorAll("span"))
      .map((span) => normalizeText(span.innerText))
      .filter(Boolean);
    const plate = texts.find((text) => /^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$/i.test(text));
    return plate || texts[0] || normalizeText(cell.innerText);
  }

  function mergeRows(left, right) {
    const merged = { ...left };
    for (const [key, value] of Object.entries(right)) {
      if (!merged[key] && value) {
        merged[key] = value;
      }
    }
    return merged;
  }

  function findSelectedStatus() {
    const labels = new Set(["aguardando revisao", "revisado", "arquivado"]);
    const buttons = Array.from(document.querySelectorAll("button"))
      .filter(isElementVisible)
      .map((button) => ({
        button,
        text: normalizeText(button.innerText)
      }))
      .filter((entry) => labels.has(normalizeForCompare(entry.text)));

    const selected = buttons
      .map((entry) => ({
        ...entry,
        score: selectedStatusButtonScore(entry.button)
      }))
      .filter((entry) => entry.score > 0)
      .sort((a, b) => b.score - a.score)[0];

    return selected?.text || buttons[0]?.text || "";
  }

  function selectedStatusButtonScore(button) {
    let score = 0;
    const classList = button.classList;
    const state = normalizeForCompare([
      button.getAttribute("aria-pressed"),
      button.getAttribute("aria-selected"),
      button.getAttribute("data-state"),
      button.getAttribute("data-active")
    ].filter(Boolean).join(" "));

    if (state.includes("true") || state.includes("active") || state.includes("selected")) {
      score += 100;
    }

    if (classList.contains("bg-darkBlue") || classList.contains("bg-primary") || classList.contains("bg-primary-950")) {
      score += 80;
    }

    if (classList.contains("text-white") && !classList.contains("bg-gray-100") && !classList.contains("bg-gray-200")) {
      score += 40;
    }

    const style = getComputedStyle(button);
    const background = parseRgbColor(style.backgroundColor);
    const color = parseRgbColor(style.color);
    if (background && color && colorLuma(background) < 95 && colorLuma(color) > 170) {
      score += 70;
    }

    return score;
  }

  function parseRgbColor(value) {
    const match = String(value || "").match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
    if (!match) {
      return null;
    }

    return {
      red: Number(match[1]),
      green: Number(match[2]),
      blue: Number(match[3])
    };
  }

  function colorLuma(color) {
    return (0.2126 * color.red) + (0.7152 * color.green) + (0.0722 * color.blue);
  }

  function findDateRangeText() {
    const candidates = Array.from(document.querySelectorAll("button, [role='button'], div, span, output"))
      .filter(isElementVisible)
      .map((element) => normalizeText(element.innerText || element.textContent || ""))
      .filter((text) => text.length <= 80);

    for (const text of candidates) {
      const match = text.match(RANGE_PATTERN);
      if (match) {
        return match[0];
      }
    }

    const bodyMatch = normalizeText(document.body.innerText).match(RANGE_PATTERN);
    return bodyMatch ? bodyMatch[0] : "";
  }

  function findDateFilterText() {
    const rangeText = findDateRangeText();
    if (rangeText) {
      return rangeText;
    }

    const trigger = findDateRangeTrigger();
    return trigger ? normalizeText(trigger.innerText || trigger.textContent || "") : "";
  }

  function findDateRangeTrigger() {
    const elements = Array.from(document.querySelectorAll("button, [role='button'], div, span"))
      .filter(isElementVisible);
    const candidates = [];

    for (const element of elements) {
      const text = normalizeText(element.innerText || element.textContent || "");
      if (text.length <= 80 && isDateFilterTriggerText(text)) {
        const target = element.closest("button, [role='button'], .group, [class*='cursor-pointer']") || element;
        if (target.getAttribute("role") === "menuitem" || target.closest("[role='menu']")) {
          continue;
        }
        candidates.push({
          element: target,
          score: RANGE_PATTERN.test(text) ? 10 : normalizeForCompare(text).includes("mes atual") ? 8 : 5,
          textLength: text.length
        });
      }
    }

    candidates.sort((a, b) => (b.score - a.score) || (a.textLength - b.textLength));
    return candidates[0]?.element || null;
  }

  function isDateFilterTriggerText(text) {
    const comparable = normalizeForCompare(text);
    return RANGE_PATTERN.test(text)
      || comparable === "mes atual"
      || comparable === "personalizado"
      || comparable.includes("ultimos")
      || comparable.includes("hoje")
      || comparable.includes("ontem")
      || comparable.includes("semana atual");
  }

  async function openCustomDatePicker() {
    if (findReactDatepickerMonthContainer()) {
      return findVisibleDateInputs();
    }

    const trigger = findDateRangeTrigger();
    if (!trigger) {
      throw new Error("Nao encontrei o seletor de periodo na tela.");
    }

    activateElement(trigger);
    await waitForRender(450);

    let customItem = findCustomDateModeItem();
    if (customItem) {
      activateElement(customItem);
      await waitForRender(450);
    }

    let inputs = findVisibleDateInputs();
    if (inputs.length >= 2) {
      return inputs;
    }

    customItem = findCustomDateModeItem();
    if (customItem) {
      customItem.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
      customItem.dispatchEvent(new MouseEvent("mouseenter", { bubbles: true }));
      await waitForRender(350);
      activateElement(customItem);
      await waitForRender(450);
    }

    return findVisibleDateInputs();
  }

  function findCustomDateModeItem() {
    return Array.from(document.querySelectorAll("[role='menuitem'], button, [role='button'], div"))
      .filter(isElementVisible)
      .find((element) => normalizeForCompare(element.innerText || element.textContent || "") === "personalizado");
  }

  function activateElement(element) {
    element.focus?.();
    element.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true }));
    element.click();
  }

  async function applyDateRange(payload = {}) {
    if (!isOccurrencesPage()) {
      throw new Error("Abra a pagina de ocorrencias antes de aplicar periodo.");
    }

    const startDate = payload.startDate;
    const endDate = payload.endDate;
    if (!isIsoDate(startDate) || !isIsoDate(endDate)) {
      throw new Error("Informe data inicial e final.");
    }

    await openCustomDatePicker();

    let applied = false;
    let inputs = findVisibleDateInputs();
    if (inputs.length >= 2) {
      setDateInputValue(inputs[0], startDate);
      await waitForRender(160);
      inputs = findVisibleDateInputs();
      setDateInputValue(inputs[1] || inputs[0], endDate);
      await waitForRender(120);
      applyOpenDatePickerSelection(inputs[1] || inputs[0]);
      applied = await waitForDateRangeApplied(startDate, endDate, 5000);
    }

    if (!applied) {
      await openCustomDatePicker();
      if (!await tryApplyDateRangeByCalendar(startDate, endDate)) {
        throw new Error("Nao consegui selecionar os dias no calendario Personalizado.");
      }
      applied = await waitForDateRangeApplied(startDate, endDate, 5000);
    }

    const currentRangeText = findDateRangeText();
    return {
      ok: applied,
      rangeText: currentRangeText,
      message: applied
        ? "Periodo aplicado."
        : `O Infleet voltou para ${currentRangeText || "outro periodo"}. A exportacao vai filtrar pelo periodo escolhido.`
    };
  }

  function findVisibleDateInputs() {
    const roots = Array.from(document.querySelectorAll("[role='dialog'], [data-radix-popper-content-wrapper], .react-datepicker, .rdp, [class*='calendar'], [class*='date']"))
      .filter(isElementVisible);
    const scope = roots.length > 0 ? roots[roots.length - 1] : document.body;
    const inputs = Array.from(scope.querySelectorAll("input"))
      .filter(isElementVisible)
      .filter((input) => {
        const text = normalizeText([
          input.type,
          input.name,
          input.id,
          input.placeholder,
          input.getAttribute("aria-label")
        ].filter(Boolean).join(" "));
        return input.type === "date"
          || input.type === "text"
          || input.type === "tel"
          || /data|date|inicio|inicial|fim|final|dd|mm|aaaa|yyyy/i.test(text);
      });

    return inputs.slice(0, 2);
  }

  function findApplyButton() {
    const pattern = /^(aplicar|filtrar|confirmar|buscar|ok|salvar)$/i;
    return Array.from(document.querySelectorAll("button"))
      .filter(isElementVisible)
      .find((button) => pattern.test(normalizeText(button.innerText)));
  }

  function applyOpenDatePickerSelection(fallbackElement) {
    const button = findApplyButton();
    if (button) {
      activateElement(button);
      return;
    }

    fallbackElement?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true }));
  }

  async function waitForDateRangeApplied(startDate, endDate, timeoutMs) {
    return waitUntil(() => {
      const rangeText = findDateRangeText();
      return dateRangeMatchesTarget(rangeText, startDate, endDate);
    }, timeoutMs);
  }

  async function tryApplyDateRangeByCalendar(startDate, endDate) {
    await openCustomDatePicker();
    await ensureReactDatepickerDailyMode();

    if (!await navigateReactDatepickerToMonth(startDate)) {
      return false;
    }

    const clickedStart = clickCalendarDate(startDate);
    if (!clickedStart) {
      return false;
    }
    await waitForRender(220);

    if (!await navigateReactDatepickerToMonth(endDate)) {
      return false;
    }

    const clickedEnd = clickCalendarDate(endDate);
    if (!clickedEnd) {
      return false;
    }

    await waitForRender(160);
    applyOpenDatePickerSelection(document.activeElement);

    return true;
  }

  function clickCalendarDate(isoDate) {
    const candidates = Array.from(document.querySelectorAll(".react-datepicker__day, button, [role='button'], [role='gridcell'], [role='option'], [data-date], [aria-label]"))
      .filter(isElementVisible)
      .map((element) => ({
        element,
        target: findClickableDateTarget(element),
        score: calendarDateMatchScore(element, isoDate)
      }))
      .filter((entry) => entry.score > 0 && entry.target)
      .sort((a, b) => b.score - a.score);

    if (candidates.length === 0) {
      return false;
    }

    candidates[0].target.click();
    return true;
  }

  function findClickableDateTarget(element) {
    return element.closest("button, [role='button']") || element;
  }

  function calendarDateMatchScore(element, isoDate) {
    if (isDisabledCalendarDate(element)) {
      return 0;
    }

    const parsedIsoDate = reactDatepickerDayIso(element);
    if (parsedIsoDate === isoDate) {
      return element.classList.contains("react-datepicker__day--outside-month") ? 24 : 35;
    }

    const attrText = normalizeText([
      element.getAttribute("data-date"),
      element.getAttribute("data-day"),
      element.getAttribute("aria-label"),
      element.getAttribute("title"),
      element.getAttribute("value")
    ].filter(Boolean).join(" ")).toLowerCase();
    const bodyText = normalizeText(element.innerText || element.textContent || "").toLowerCase();
    const tokens = calendarDateTokens(isoDate).map((token) => token.toLowerCase());

    if (tokens.some((token) => attrText.includes(token))) {
      return 10;
    }

    const [year, month, day] = isoDate.split("-");
    const monthTokens = [monthNamePt(Number(month)), monthNameEn(Number(month)), month].map((token) => token.toLowerCase());
    const hasMonthAndYear = monthTokens.some((token) => attrText.includes(token)) && attrText.includes(year);
    const dayNumber = String(Number(day));

    if (hasMonthAndYear && (bodyText === day || bodyText === dayNumber || attrText.includes(` ${dayNumber} `))) {
      return 5;
    }

    return 0;
  }

  async function ensureReactDatepickerDailyMode() {
    const container = findReactDatepickerMonthContainer();
    const input = container?.querySelector("input[type='checkbox']");
    if (input?.checked) {
      activateElement(input.closest("label") || input);
      await waitForRender(250);
    }
  }

  async function navigateReactDatepickerToMonth(isoDate) {
    const target = yearMonthIndexFromIso(isoDate);
    for (let attempt = 0; attempt < 36; attempt += 1) {
      const current = currentReactDatepickerYearMonthIndex();
      if (current === target) {
        return true;
      }
      if (!Number.isFinite(current)) {
        return false;
      }

      const direction = current < target ? "next" : "prev";
      const button = findReactDatepickerNavButton(direction);
      if (!button) {
        return false;
      }

      activateElement(button);
      await waitForRender(280);
    }

    return currentReactDatepickerYearMonthIndex() === target;
  }

  function currentReactDatepickerYearMonthIndex() {
    const month = findReactDatepickerMonthContainer()?.querySelector(".react-datepicker__month[aria-label*='month']");
    const match = String(month?.getAttribute("aria-label") || "").match(/(\d{4})-(\d{2})/);
    if (!match) {
      return Number.NaN;
    }

    return (Number(match[1]) * 12) + Number(match[2]);
  }

  function yearMonthIndexFromIso(isoDate) {
    const [year, month] = isoDate.split("-");
    return (Number(year) * 12) + Number(month);
  }

  function findReactDatepickerMonthContainer() {
    return Array.from(document.querySelectorAll(".react-datepicker__month-container"))
      .filter(isElementVisible)[0] || null;
  }

  function findReactDatepickerNavButton(direction) {
    const container = findReactDatepickerMonthContainer();
    const header = container?.querySelector(".react-datepicker__header") || container;
    const buttons = Array.from(header?.querySelectorAll("button") || [])
      .filter(isElementVisible);

    if (buttons.length < 2) {
      return null;
    }

    return direction === "prev" ? buttons[0] : buttons[buttons.length - 1];
  }

  function reactDatepickerDayIso(element) {
    const text = normalizeText(element.getAttribute("aria-label") || "");
    const match = text.match(/(?:Choose\s+)?(?:[A-Za-z]+,\s+)?([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?,\s+(\d{4})/i);
    if (!match) {
      return "";
    }

    const month = monthNumberEn(match[1]);
    if (!month) {
      return "";
    }

    return `${match[3]}-${String(month).padStart(2, "0")}-${String(Number(match[2])).padStart(2, "0")}`;
  }

  function isDisabledCalendarDate(element) {
    return element.getAttribute("aria-disabled") === "true"
      || element.classList.contains("react-datepicker__day--disabled")
      || element.hasAttribute("disabled");
  }

  function calendarDateTokens(isoDate) {
    const [year, month, day] = isoDate.split("-");
    const dayNumber = String(Number(day));
    const ptMonth = monthNamePt(Number(month));
    const enMonth = monthNameEn(Number(month));

    return [
      isoDate,
      isoToBrazil(isoDate, false),
      isoToBrazil(isoDate, true),
      `${dayNumber} de ${ptMonth} de ${year}`,
      `${dayNumber} ${ptMonth} ${year}`,
      `${enMonth} ${dayNumber}, ${year}`,
      `${enMonth} ${dayNumber}${englishOrdinalSuffix(Number(dayNumber))}, ${year}`,
      `${dayNumber} ${enMonth} ${year}`
    ];
  }

  function monthNamePt(month) {
    return [
      "janeiro",
      "fevereiro",
      "marco",
      "abril",
      "maio",
      "junho",
      "julho",
      "agosto",
      "setembro",
      "outubro",
      "novembro",
      "dezembro"
    ][month - 1] || "";
  }

  function monthNameEn(month) {
    return [
      "january",
      "february",
      "march",
      "april",
      "may",
      "june",
      "july",
      "august",
      "september",
      "october",
      "november",
      "december"
    ][month - 1] || "";
  }

  function monthNumberEn(monthName) {
    return [
      "january",
      "february",
      "march",
      "april",
      "may",
      "june",
      "july",
      "august",
      "september",
      "october",
      "november",
      "december"
    ].indexOf(String(monthName || "").toLowerCase()) + 1;
  }

  function englishOrdinalSuffix(day) {
    if (day >= 11 && day <= 13) {
      return "th";
    }

    return { 1: "st", 2: "nd", 3: "rd" }[day % 10] || "th";
  }

  function setDateInputValue(input, isoDate) {
    const candidates = inputValueCandidates(input, isoDate);
    for (const value of candidates) {
      setInputValue(input, value);
      if (input.type === "date" || normalizeText(input.value) === value) {
        return;
      }
    }
  }

  function setInputValue(input, value) {
    input.focus();
    if (typeof input.select === "function") {
      input.select();
    }
    const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
    if (descriptor && typeof descriptor.set === "function") {
      descriptor.set.call(input, "");
      input.dispatchEvent(new Event("input", { bubbles: true }));
      descriptor.set.call(input, value);
    } else {
      input.value = "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.value = value;
    }
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent("keyup", { key: "Tab", bubbles: true }));
    input.blur();
  }

  function inputValueCandidates(input, isoDate) {
    if (input.type === "date") {
      return [isoDate];
    }

    const placeholder = normalizeText([
      input.placeholder,
      input.getAttribute("aria-label"),
      input.getAttribute("name"),
      input.getAttribute("id")
    ].filter(Boolean).join(" ")).toLowerCase();
    const maxLength = Number(input.getAttribute("maxlength") || input.maxLength || 0);
    const shortValue = isoToBrazil(isoDate, true);
    const longValue = isoToBrazil(isoDate, false);

    if ((maxLength > 0 && maxLength <= 8) || /\baa\b|\byy\b/.test(placeholder)) {
      return [shortValue, longValue];
    }

    return [longValue, shortValue];
  }

  function findScrollContainer() {
    const links = Array.from(document.querySelectorAll(OCCURRENCE_LINK_SELECTOR));
    if (links.length === 0) {
      return document.scrollingElement || document.documentElement;
    }

    const candidates = Array.from(document.querySelectorAll("body, body *"))
      .filter((element) => {
        if (!element.contains(links[0])) {
          return false;
        }
        const style = getComputedStyle(element);
        const overflowY = style.overflowY;
        return (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay")
          && element.scrollHeight > element.clientHeight + 80;
      })
      .sort((a, b) => (a.clientHeight - b.clientHeight) || (a.scrollHeight - b.scrollHeight));

    return candidates[0] || document.scrollingElement || document.documentElement;
  }

  function getScrollTop(container) {
    return isDocumentScroller(container) ? window.scrollY : container.scrollTop;
  }

  function getScrollHeight(container) {
    return isDocumentScroller(container)
      ? Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)
      : container.scrollHeight;
  }

  function getClientHeight(container) {
    return isDocumentScroller(container) ? window.innerHeight : container.clientHeight;
  }

  async function setScrollTop(container, value) {
    if (isDocumentScroller(container)) {
      window.scrollTo({ top: value, behavior: "auto" });
    } else {
      container.scrollTop = value;
      container.dispatchEvent(new Event("scroll", { bubbles: true }));
    }
    await waitForRender(40);
  }

  function isDocumentScroller(container) {
    return container === document.body
      || container === document.documentElement
      || container === document.scrollingElement;
  }

  function downloadBytes(bytes, filename) {
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function makeFilename(startDate, endDate, rangeText) {
    const safeRange = startDate && endDate
      ? `${startDate}_a_${endDate}`
      : String(rangeText || "periodo")
        .replace(/\s*-\s*/g, "_a_")
        .replace(/\//g, "-")
        .replace(/[^\w.-]+/g, "_");
    return `infleet-ocorrencias_${safeRange}.xlsx`;
  }

  function filterRowsByRequestedRange(rows, startDate, endDate) {
    if (!isIsoDate(startDate) || !isIsoDate(endDate)) {
      return rows;
    }

    return rows.filter((row) => {
      const rowDate = brazilDateToIso(row.date);
      return rowDate && rowDate >= startDate && rowDate <= endDate;
    });
  }

  function dateRangeMatchesTarget(rangeText, startDate, endDate) {
    const dates = parseDateRangeText(rangeText);
    return dates
      && dates.startDate === startDate
      && dates.endDate === endDate;
  }

  function parseDateRangeText(rangeText) {
    const match = String(rangeText || "").match(/(\d{2})\/(\d{2})\/(\d{2}|\d{4})\s*-\s*(\d{2})\/(\d{2})\/(\d{2}|\d{4})/);
    if (!match) {
      return null;
    }

    return {
      startDate: brazilPartsToIso(match[1], match[2], match[3]),
      endDate: brazilPartsToIso(match[4], match[5], match[6])
    };
  }

  function brazilDateToIso(dateText) {
    const match = String(dateText || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return match ? brazilPartsToIso(match[1], match[2], match[3]) : "";
  }

  function brazilPartsToIso(day, month, year) {
    const fullYear = year.length === 2 ? `20${year}` : year;
    return `${fullYear}-${month}-${day}`;
  }

  function formatRequestedRange(startDate, endDate) {
    if (!isIsoDate(startDate) || !isIsoDate(endDate)) {
      return "";
    }
    return `${isoToBrazil(startDate, true)} - ${isoToBrazil(endDate, true)}`;
  }

  function formatDateTime(date) {
    const pad = (value) => String(value).padStart(2, "0");
    return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
  }

  function isoToBrazil(isoDate, shortYear) {
    const [year, month, day] = isoDate.split("-");
    return `${day}/${month}/${shortYear ? year.slice(-2) : year}`;
  }

  function isIsoDate(value) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""));
  }

  function normalizeText(value) {
    return String(value ?? "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeForCompare(value) {
    return normalizeText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function isElementVisible(element) {
    if (!element || !(element instanceof Element)) {
      return false;
    }
    const style = getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden" || Number(style.opacity) === 0) {
      return false;
    }
    const rect = element.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function waitForRender(ms) {
    return new Promise((resolve) => {
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          setTimeout(resolve, ms);
        });
      });
    });
  }

  async function waitUntil(predicate, timeoutMs) {
    const started = Date.now();
    while (Date.now() - started < timeoutMs) {
      if (predicate()) {
        return true;
      }
      await waitForRender(150);
    }
    return false;
  }

  function sendProgress(payload) {
    try {
      const result = chrome.runtime.sendMessage({ type: MESSAGE.PROGRESS, payload });
      if (result && typeof result.catch === "function") {
        result.catch(() => {});
      }
    } catch {
      // Popup fechado ou runtime indisponivel: a exportacao continua pela tela.
    }
  }

  function updateOverlay(message, state) {
    const overlay = ensureOverlay(state === "running");
    const title = overlay.querySelector("[data-infleet-overlay-title]");
    const detail = overlay.querySelector("[data-infleet-overlay-detail]");

    title.textContent = state === "running"
      ? "Extração em andamento"
      : state === "success"
        ? "Extração concluída"
        : state === "error"
          ? "Falha na extração"
          : "Extração";
    detail.textContent = message;
    overlay.dataset.state = state;
    overlay.style.pointerEvents = state === "running" ? "all" : "none";

    if (state === "running") {
      lockPageInteraction();
    } else {
      unlockPageInteraction();
    }

    if (state === "success" || state === "error" || state === "idle") {
      setTimeout(() => {
        if (overlay.dataset.state === state) {
          overlay.remove();
        }
      }, state === "error" ? 9000 : 5500);
    }
  }

  function ensureOverlay(blocking) {
    const id = "infleet-occurrence-exporter-overlay";
    let overlay = document.getElementById(id);
    if (overlay) {
      overlay.style.pointerEvents = blocking ? "all" : "none";
      return overlay;
    }

    overlay = document.createElement("div");
    overlay.id = id;
    overlay.setAttribute("role", "status");
    overlay.style.position = "fixed";
    overlay.style.inset = "0";
    overlay.style.zIndex = "2147483647";
    overlay.style.display = "flex";
    overlay.style.alignItems = "center";
    overlay.style.justifyContent = "center";
    overlay.style.background = "rgba(15, 23, 42, 0.38)";
    overlay.style.backdropFilter = "blur(2px)";
    overlay.style.color = "#172033";
    overlay.style.font = "14px/1.45 Arial, Helvetica, sans-serif";
    overlay.style.letterSpacing = "0";
    overlay.style.pointerEvents = blocking ? "all" : "none";

    const panel = document.createElement("div");
    panel.style.width = "min(420px, calc(100vw - 32px))";
    panel.style.border = "1px solid rgba(148, 163, 184, 0.45)";
    panel.style.borderRadius = "8px";
    panel.style.background = "#ffffff";
    panel.style.boxShadow = "0 22px 60px rgba(15, 23, 42, 0.32)";
    panel.style.padding = "20px";
    panel.style.textAlign = "center";

    const spinner = document.createElement("div");
    spinner.style.width = "34px";
    spinner.style.height = "34px";
    spinner.style.margin = "0 auto 14px";
    spinner.style.border = "4px solid #dbe4f0";
    spinner.style.borderTopColor = "#153c78";
    spinner.style.borderRadius = "50%";
    spinner.style.animation = "infleet-export-spin 0.9s linear infinite";

    const title = document.createElement("strong");
    title.dataset.inflight = "title";
    title.setAttribute("data-infleet-overlay-title", "");
    title.style.display = "block";
    title.style.marginBottom = "6px";
    title.style.fontSize = "17px";
    title.style.lineHeight = "1.25";

    const detail = document.createElement("div");
    detail.setAttribute("data-infleet-overlay-detail", "");
    detail.style.color = "#526071";
    detail.style.fontSize = "13px";
    detail.style.minHeight = "19px";

    const note = document.createElement("div");
    note.textContent = "Aguarde.";
    note.style.marginTop = "12px";
    note.style.color = "#7c8797";
    note.style.fontSize = "12px";

    const style = document.createElement("style");
    style.textContent = "@keyframes infleet-export-spin { to { transform: rotate(360deg); } }";

    panel.append(spinner, title, detail, note);
    overlay.append(style, panel);
    document.body.appendChild(overlay);
    return overlay;
  }

  function lockPageInteraction() {
    if (overlayCleanup) {
      return;
    }

    const prevent = (event) => {
      event.preventDefault();
      event.stopPropagation();
    };
    const preventScrollKey = (event) => {
      const keys = ["ArrowDown", "ArrowUp", "PageDown", "PageUp", "Home", "End", " "];
      if (keys.includes(event.key)) {
        prevent(event);
      }
    };
    const options = { capture: true, passive: false };

    document.addEventListener("wheel", prevent, options);
    document.addEventListener("touchmove", prevent, options);
    document.addEventListener("keydown", preventScrollKey, true);

    overlayCleanup = () => {
      document.removeEventListener("wheel", prevent, options);
      document.removeEventListener("touchmove", prevent, options);
      document.removeEventListener("keydown", preventScrollKey, true);
      overlayCleanup = null;
    };
  }

  function unlockPageInteraction() {
    if (overlayCleanup) {
      overlayCleanup();
    }
  }
})();
