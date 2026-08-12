(function initPopup() {
  "use strict";

  const MESSAGE = {
    PAGE_INFO: "INFLEET_PAGE_INFO",
    EXPORT: "INFLEET_EXPORT",
    APPLY_DATE: "INFLEET_APPLY_DATE_RANGE",
    CANCEL: "INFLEET_CANCEL_EXPORT",
    PROGRESS: "INFLEET_EXPORT_PROGRESS"
  };

  const elements = {
    pageStatus: document.getElementById("pageStatus"),
    refreshButton: document.getElementById("refreshButton"),
    applyDateButton: document.getElementById("applyDateButton"),
    exportButton: document.getElementById("exportButton"),
    startDate: document.getElementById("startDate"),
    endDate: document.getElementById("endDate"),
    rangeText: document.getElementById("rangeText"),
    visibleRows: document.getElementById("visibleRows"),
    statusText: document.getElementById("statusText"),
    scrollAll: document.getElementById("scrollAll"),
    message: document.getElementById("message")
  };

  let activeTabId = null;
  let exporting = false;

  document.addEventListener("DOMContentLoaded", refreshInfo);
  elements.refreshButton.addEventListener("click", refreshInfo);
  elements.applyDateButton.addEventListener("click", applyDateRange);
  elements.exportButton.addEventListener("click", exportExcel);

  chrome.runtime.onMessage.addListener((request) => {
    if (!request || request.type !== MESSAGE.PROGRESS) {
      return;
    }
    const payload = request.payload || {};
    setMessage(payload.message || "", payload.phase === "error" ? "error" : payload.phase === "done" ? "ok" : "");
    if (payload.rows != null) {
      elements.visibleRows.textContent = String(payload.rows);
    }
    if (payload.phase === "done" || payload.phase === "error") {
      setBusy(false);
      exporting = false;
      refreshInfo();
    }
  });

  async function refreshInfo() {
    try {
      const tab = await getActiveTab();
      activeTabId = tab.id;
      if (!isInfleetTab(tab)) {
        renderUnavailable("Abra a tela /occurrences do Infleet.");
        return;
      }

      const response = await sendToTab(tab.id, { type: MESSAGE.PAGE_INFO });
      if (!response?.ok) {
        renderUnavailable(response?.message || "Recarregue a página do Infleet.");
        return;
      }

      elements.pageStatus.textContent = response.isOccurrencesPage
        ? "Página de ocorrências conectada."
        : "Você está no Infleet, mas fora de /occurrences.";
      elements.rangeText.textContent = response.rangeText || response.dateFilterText || "-";
      elements.visibleRows.textContent = String(response.visibleRows ?? 0);
      elements.statusText.textContent = response.selectedStatus || "-";
      parseRangeIntoInputs(response.rangeText);
      setControlsEnabled(response.isOccurrencesPage);
      setMessage(response.scrollable ? "" : "Lista sem rolagem detectada.", "");
    } catch (error) {
      renderUnavailable(error.message);
    }
  }

  async function applyDateRange() {
    try {
      const payload = getDatePayload();
      if (!payload) {
        return;
      }
      setBusy(true);
      setMessage("Aplicando período no Infleet...", "");
      const response = await sendToActiveTab({ type: MESSAGE.APPLY_DATE, payload });
      await refreshInfo();
      elements.startDate.value = payload.startDate;
      elements.endDate.value = payload.endDate;
      setMessage(response.message || "Período enviado.", response.ok ? "ok" : "");
    } catch (error) {
      setMessage(error.message, "error");
    } finally {
      if (!exporting) {
        setBusy(false);
      }
    }
  }

  async function exportExcel() {
    try {
      const payload = {
        ...getDatePayload(false),
        scrollAll: elements.scrollAll.checked,
        restoreScroll: true
      };
      exporting = true;
      setBusy(true);
      setMessage("Exportação iniciada. Acompanhe o progresso na página.", "");
      const response = await sendToActiveTab({ type: MESSAGE.EXPORT, payload });
      await refreshInfo();
      if (payload.startDate && payload.endDate) {
        elements.startDate.value = payload.startDate;
        elements.endDate.value = payload.endDate;
      }
      if (!response.ok) {
        setMessage(response.message || "Exportação não concluída.", response.cancelled ? "" : "error");
      } else {
        setMessage(`${response.rows} ocorrências exportadas.`, "ok");
      }
    } catch (error) {
      setMessage(error.message, "error");
    } finally {
      exporting = false;
      setBusy(false);
    }
  }

  function getDatePayload(required = true) {
    const startDate = elements.startDate.value;
    const endDate = elements.endDate.value;
    if (!startDate && !endDate && !required) {
      return {};
    }
    if (!startDate || !endDate) {
      setMessage("Informe a data inicial e final.", "error");
      return null;
    }
    if (startDate > endDate) {
      setMessage("A data inicial não pode ser maior que a final.", "error");
      return null;
    }
    return { startDate, endDate };
  }

  async function sendToActiveTab(message) {
    const tab = await getActiveTab();
    activeTabId = tab.id;
    if (!isInfleetTab(tab)) {
      throw new Error("Abra https://app.infleet.com.br/occurrences.");
    }
    return sendToTab(tab.id, message);
  }

  async function getActiveTab() {
    return new Promise((resolve, reject) => {
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        const error = chrome.runtime.lastError;
        if (error) {
          reject(new Error(error.message));
          return;
        }
        const tab = tabs && tabs[0];
        if (!tab || !tab.id) {
          reject(new Error("Nenhuma aba ativa encontrada."));
          return;
        }
        resolve(tab);
      });
    });
  }

  function sendToTab(tabId, message) {
    return new Promise((resolve, reject) => {
      chrome.tabs.sendMessage(tabId, message, (response) => {
        const error = chrome.runtime.lastError;
        if (error) {
          reject(new Error("Recarregue a página da Infleet depois de instalar a extensão."));
          return;
        }
        resolve(response);
      });
    });
  }

  function isInfleetTab(tab) {
    return /^https:\/\/app\.infleet\.com\.br\//.test(tab.url || "");
  }

  function renderUnavailable(message) {
    elements.pageStatus.textContent = message;
    elements.rangeText.textContent = "-";
    elements.visibleRows.textContent = "0";
    elements.statusText.textContent = "-";
    setControlsEnabled(false);
    setMessage(message, "error");
  }

  function setBusy(isBusy) {
    elements.refreshButton.disabled = isBusy;
    elements.applyDateButton.disabled = isBusy;
    elements.exportButton.disabled = isBusy;
  }

  function setControlsEnabled(enabled) {
    elements.applyDateButton.disabled = !enabled;
    elements.exportButton.disabled = !enabled;
  }

  function setMessage(message, type) {
    elements.message.textContent = message || "";
    elements.message.className = `message${type ? ` ${type}` : ""}`;
  }

  function parseRangeIntoInputs(rangeText) {
    const match = String(rangeText || "").match(/(\d{2})\/(\d{2})\/(\d{2}|\d{4})\s*-\s*(\d{2})\/(\d{2})\/(\d{2}|\d{4})/);
    if (!match) {
      return;
    }
    const start = brazilToIso(match[1], match[2], match[3]);
    const end = brazilToIso(match[4], match[5], match[6]);
    if (start) {
      elements.startDate.value = start;
    }
    if (end) {
      elements.endDate.value = end;
    }
  }

  function brazilToIso(day, month, year) {
    const fullYear = year.length === 2 ? `20${year}` : year;
    return `${fullYear}-${month}-${day}`;
  }
})();
