// static/js/main.js

// Листы Excel
const COMPETITORS_SHEET_NAME = "Конкуренты";
const NPZ_SHEET_NAME = "НПЗ";
const PETROPAVLOVSK_SHEET_NAME = "Петропавловск - Камчатский";
const HIDDEN_SHEET_NAMES = ["Курсы"]; // не показываем отдельным блоком

// Сохранение порядка блоков
const ORDER_STORAGE_KEY = "dashboard_block_order_v1";

// Архив дат
let availableDates = [];
let activeDate = null; // строка вида "2025-12-08"

document.addEventListener("DOMContentLoaded", () => {
  setupArchiveUI();
  fetchDatesAndInit();

  const screenshotBtn = document.getElementById("screenshotButton");
  if (screenshotBtn) {
    screenshotBtn.addEventListener("click", handleScreenshotClick);
  }
});

/* ---------------------- СКРИНШОТ ---------------------- */

async function handleScreenshotClick() {
  if (typeof html2canvas === "undefined") {
    alert("Скриншот недоступен: html2canvas не загружен.");
    return;
  }

  const target = document.querySelector(".page") || document.body;

  try {
    const canvas = await html2canvas(target, {
      windowWidth: document.documentElement.scrollWidth,
      windowHeight: document.documentElement.scrollHeight,
    });

    canvas.toBlob((blob) => {
      if (!blob) return;

      const link = document.createElement("a");

      const now = new Date();
      const pad = (n) => String(n).padStart(2, "0");
      const fileName = `trastboard_${now.getFullYear()}-${pad(
        now.getMonth() + 1
      )}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(
        now.getMinutes()
      )}-${pad(now.getSeconds())}.png`;

      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);
    });
  } catch (err) {
    console.error("Ошибка при создании скрина:", err);
  }
}

/* ---------------------- АРХИВ ДАТ ---------------------- */

function setupArchiveUI() {
  const btn = document.getElementById("archiveButton");
  const menu = document.getElementById("archiveMenu");

  if (!btn || !menu) return;

  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const isHidden = menu.hasAttribute("hidden");
    if (isHidden) {
      menu.removeAttribute("hidden");
    } else {
      menu.setAttribute("hidden", "true");
    }
  });

  // Кликом вне меню закрываем его
  document.addEventListener("click", () => {
    if (!menu.hasAttribute("hidden")) {
      menu.setAttribute("hidden", "true");
    }
  });

  // Не даём клику внутри меню всплывать до документа
  menu.addEventListener("click", (e) => e.stopPropagation());
}

async function fetchDatesAndInit() {
  try {
    const res = await fetch("/api/dates");
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Ошибка загрузки дат");
    }

    availableDates = Array.isArray(data.dates) ? data.dates : [];

    if (availableDates.length > 0) {
      activeDate = availableDates[0]; // последняя дата
      updateActiveDateLabel();
      renderArchiveMenu();
      await fetchBlocksForActiveDate();
    } else {
      // дат нет — работаем по режиму "последний день по листам"
      activeDate = null;
      updateActiveDateLabel();
      renderArchiveMenu();
      await fetchBlocks(null);
    }
  } catch (err) {
    console.error("Ошибка при инициализации дат:", err);
    // Если не смогли загрузить даты, пробуем просто блоки
    await fetchBlocks(null);
  }
}

function updateActiveDateLabel() {
  const label = document.getElementById("activeDateLabel");
  if (!label) return;

  if (!activeDate) {
    label.textContent = "—";
  } else {
    label.textContent = formatDateRu(activeDate);
  }
}

function renderArchiveMenu() {
  const menu = document.getElementById("archiveMenu");
  if (!menu) return;

  menu.innerHTML = "";

  if (!availableDates.length) {
    const empty = document.createElement("div");
    empty.textContent = "Нет дат";
    empty.style.fontSize = "12px";
    empty.style.color = "#6b7280";
    empty.style.padding = "6px 8px";
    menu.appendChild(empty);
    return;
  }

  availableDates.forEach((dateStr) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "archive-menu__item";
    btn.textContent = formatDateRu(dateStr);
    btn.dataset.date = dateStr;

    if (dateStr === activeDate) {
      btn.classList.add("is-active");
    }

    btn.addEventListener("click", async () => {
      activeDate = dateStr;
      updateActiveDateLabel();
      renderArchiveMenu(); // перерисуем, чтобы подсветить активную дату
      await fetchBlocksForActiveDate();

      const menuEl = document.getElementById("archiveMenu");
      if (menuEl) {
        menuEl.setAttribute("hidden", "true");
      }
    });

    menu.appendChild(btn);
  });
}

async function fetchBlocksForActiveDate() {
  if (!activeDate) {
    return fetchBlocks(null);
  }
  return fetchBlocks(activeDate);
}

/* ---------------------- ЗАГРУЗКА БЛОКОВ ---------------------- */

async function fetchBlocks(dateStr) {
  let url = "/api/blocks";
  if (dateStr) {
    url += "?date=" + encodeURIComponent(dateStr);
  }

  try {
    const res = await fetch(url);
    const data = await res.json();

    if (!res.ok) {
      const msg = data && data.error ? data.error : "Ошибка загрузки данных.";
      showError(msg);
      return;
    }

    if (!data || !Array.isArray(data.blocks)) {
      showError("Некорректный формат ответа от сервера.");
      return;
    }

    renderBlocks(data.blocks);
  } catch (err) {
    console.error(err);
    showError(
      "Не удалось загрузить данные. Проверь, что сервер запущен (python app.py) и Excel доступен."
    );
  }
}

function showError(message) {
  const errorNode = document.getElementById("errorContainer");
  if (!errorNode) return;

  errorNode.textContent = message;
  errorNode.hidden = false;
}

function renderBlocks(blocks) {
  const container = document.getElementById("blocksContainer");
  if (!container) return;

  const visibleBlocks = blocks.filter(
    (b) => !HIDDEN_SHEET_NAMES.includes(b.sheetName)
  );

  const ordered = applySavedOrder(visibleBlocks);

  container.innerHTML = "";

  ordered.forEach((block) => {
    const card = createBlockCard(block);
    container.appendChild(card);
  });

  initDragDrop(container);
}

function applySavedOrder(blocks) {
  const raw = localStorage.getItem(ORDER_STORAGE_KEY);
  if (!raw) return blocks;

  let savedOrder;
  try {
    savedOrder = JSON.parse(raw);
  } catch {
    return blocks;
  }

  if (!Array.isArray(savedOrder) || !savedOrder.length) {
    return blocks;
  }

  const map = new Map();
  blocks.forEach((b) => map.set(String(b.id), b));

  const result = [];

  savedOrder.forEach((id) => {
    const key = String(id);
    if (map.has(key)) {
      result.push(map.get(key));
      map.delete(key);
    }
  });

  for (const [, block] of map) {
    result.push(block);
  }

  return result;
}

/* ---------------------- КАРТОЧКИ БЛОКОВ ---------------------- */

function createBlockCard(block) {
  const isCompetitors = block.sheetName === COMPETITORS_SHEET_NAME;
  const isNPZ = block.sheetName === NPZ_SHEET_NAME;
  const isPetropavlovsk = block.sheetName === PETROPAVLOVSK_SHEET_NAME;

  const card = document.createElement("section");
  card.className = "block-card";
  card.setAttribute("draggable", "true");
  card.dataset.blockId = String(block.id);

  if (isCompetitors) {
    card.classList.add("block-card--full", "block-card--competitors");
  } else if (isNPZ) {
    card.classList.add("block-card--full");
  }

  const header = document.createElement("div");
  header.className = "block-header";

  const headerMain = document.createElement("div");
  headerMain.className = "block-header-main";

  const title = document.createElement("h2");
  title.className = "block-title";

  if (isCompetitors) {
    title.textContent = "Конкурентные цены";
  } else if (isNPZ) {
    title.textContent = "НПЗ";
  } else if (isPetropavlovsk) {
    title.textContent = "Петропавловск-Камчатский";
  } else {
    title.textContent = block.title || block.sheetName || "Без названия";
  }

  const subtitle = document.createElement("p");
  subtitle.className = "block-subtitle";
  subtitle.textContent = `Лист Excel: ${block.sheetName || "—"}`;

  headerMain.appendChild(title);
  headerMain.appendChild(subtitle);

  const badge = document.createElement("span");
  badge.className = "block-badge";
  badge.textContent = isCompetitors
    ? "Конкуренты"
    : isNPZ
    ? "НПЗ"
    : isPetropavlovsk
    ? "П-Камчатский"
    : "Блок";

  header.appendChild(headerMain);
  header.appendChild(badge);
  card.appendChild(header);

  const hideDateCol = isCompetitors || isNPZ || isPetropavlovsk;
  const customSort = isCompetitors ? competitorsRowComparator : null;

  renderTableBlock(card, block, {
    hideDateColumn: hideDateCol,
    customSort,
  });

  return card;
}

function renderTableBlock(card, block, options = {}) {
  const { hideDateColumn = false, customSort = null } = options;

  const wrapper = document.createElement("div");
  wrapper.className = "block-table-wrapper";

  const rows = block.rows || [];
  const cols = block.columns || [];

  if (!rows.length || !cols.length) {
    const empty = document.createElement("div");
    empty.className = "block-empty";
    empty.textContent =
      "Нет данных на этом листе (или нет строк для выбранной даты).";
    wrapper.appendChild(empty);
    card.appendChild(wrapper);
    return;
  }

  let visibleColIndices = cols.map((_, idx) => idx);

  if (hideDateColumn) {
    const dateIdx = findDateColumnIndex(block);
    if (dateIdx !== null && dateIdx >= 0) {
      visibleColIndices = visibleColIndices.filter((i) => i !== dateIdx);
    }
  }

  if (block.sheetName === COMPETITORS_SHEET_NAME) {
    visibleColIndices = reorderCompetitorsColumns(cols, visibleColIndices);
  }

  let workRows = [...rows];
  if (typeof customSort === "function") {
    workRows.sort(customSort);
  }

  const table = document.createElement("table");
  table.className = "block-table";

  const thead = document.createElement("thead");
  const headTr = document.createElement("tr");

  visibleColIndices.forEach((colIndex) => {
    const th = document.createElement("th");
    th.textContent = cols[colIndex] || "";
    headTr.appendChild(th);
  });

  thead.appendChild(headTr);

  const tbody = document.createElement("tbody");

  const numericOriginal = Array.isArray(block.numericColumns)
    ? block.numericColumns.map((i) => Number(i))
    : [];
  const numericNew = [];

  visibleColIndices.forEach((origIdx, newIdx) => {
    if (numericOriginal.includes(origIdx)) {
      numericNew.push(newIdx);
    }
  });

  workRows.forEach((row) => {
    const tr = document.createElement("tr");

    visibleColIndices.forEach((origIdx, newIdx) => {
      const td = document.createElement("td");
      const rawVal = row[origIdx];

      let display = "";
      if (rawVal !== null && rawVal !== undefined) {
        display = String(rawVal);
        if (numericNew.includes(newIdx)) {
          display = formatNumberWithSpaces(display);
        }
      }

      td.textContent = display;

      if (newIdx === 0) {
        td.classList.add("cell-primary");
      }

      if (numericNew.includes(newIdx)) {
        td.classList.add("cell-number");
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  wrapper.appendChild(table);
  card.appendChild(wrapper);

  enableTableSorting(table, numericNew);
  initColumnDrag(table);
}

/* ---------------------- ЛОГИКА КОНКУРЕНТОВ ---------------------- */

function reorderCompetitorsColumns(columns, visibleColIndices) {
  let productIdx = null;
  let birjaIdx = null;
  let nnkIdx = null;

  visibleColIndices.forEach((idx) => {
    const name = (columns[idx] || "").toString().trim().toLowerCase();
    if (
      productIdx === null &&
      (name.includes("продукт") || name.includes("номенклат"))
    ) {
      productIdx = idx;
    }
    if (birjaIdx === null && name.includes("биржа")) {
      birjaIdx = idx;
    }
    if (nnkIdx === null && name.includes("ннк")) {
      nnkIdx = idx;
    }
  });

  const result = [];
  const used = new Set();

  if (productIdx !== null && visibleColIndices.includes(productIdx)) {
    result.push(productIdx);
    used.add(productIdx);
  }

  if (birjaIdx !== null && visibleColIndices.includes(birjaIdx)) {
    result.push(birjaIdx);
    used.add(birjaIdx);
  }

  if (nnkIdx !== null && visibleColIndices.includes(nnkIdx)) {
    result.push(nnkIdx);
    used.add(nnkIdx);
  }

  visibleColIndices.forEach((idx) => {
    if (!used.has(idx)) {
      result.push(idx);
    }
  });

  return result;
}

function competitorsRowComparator(rowA, rowB) {
  const getName = (row) => {
    const v = row[0];
    return v === null || v === undefined ? "" : String(v).trim();
  };

  const nameA = getName(rowA);
  const nameB = getName(rowB);

  const weight = (name) => {
    const lower = name.toLowerCase();
    if (lower.includes("биржа")) return 0;
    if (lower.includes("ннк")) return 1;
    return 2;
  };

  const wA = weight(nameA);
  const wB = weight(nameB);

  if (wA !== wB) return wA - wB;
  return nameA.localeCompare(nameB);
}

/* ---------------------- ПОИСК КОЛОНКИ ДАТЫ ---------------------- */

function findDateColumnIndex(block) {
  const cols = block.columns || [];
  const rows = block.rows || [];

  if (!cols.length || !rows.length) return null;

  for (let i = 0; i < cols.length; i++) {
    const name = String(cols[i] || "").toLowerCase();
    if (name.includes("дата")) {
      return i;
    }
  }

  const isDateString = (value) => {
    if (!value) return false;
    const s = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return true;
    if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) return true;
    if (/^\d{2}-\d{2}-\d{4}$/.test(s)) return true;
    return false;
  };

  for (let colIdx = 0; colIdx < cols.length; colIdx++) {
    let hasDate = false;
    for (let r = 0; r < rows.length; r++) {
      const cell = rows[r][colIdx];
      if (isDateString(cell)) {
        hasDate = true;
        break;
      }
    }
    if (hasDate) return colIdx;
  }

  return null;
}

/* ---------------------- ПЕРЕТАСКИВАНИЕ БЛОКОВ ---------------------- */

function initDragDrop(container) {
  let draggedCard = null;

  container.querySelectorAll(".block-card").forEach((card) => {
    card.addEventListener("dragstart", (event) => {
      if (event.target.tagName === "TH") {
        return;
      }

      draggedCard = card;
      card.classList.add("dragging");
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", card.dataset.blockId || "");
    });

    card.addEventListener("dragend", () => {
      if (draggedCard) {
        draggedCard.classList.remove("dragging");
        draggedCard = null;
      }
      saveCurrentOrder(container);
    });
  });

  container.addEventListener("dragover", (event) => {
    event.preventDefault();
    const dragging = container.querySelector(".block-card.dragging");
    if (!dragging) return;

    const afterElement = getDragAfterElement(container, event.clientY);
    if (afterElement == null) {
      container.appendChild(dragging);
    } else {
      container.insertBefore(dragging, afterElement);
    }
  });
}

function getDragAfterElement(container, mouseY) {
  const cards = [
    ...container.querySelectorAll(".block-card:not(.dragging)"),
  ];

  let closest = {
    offset: Number.NEGATIVE_INFINITY,
    element: null,
  };

  cards.forEach((card) => {
    const box = card.getBoundingClientRect();
    const offset = mouseY - box.top - box.height / 2;

    if (offset < 0 && offset > closest.offset) {
      closest = { offset, element: card };
    }
  });

  return closest.element;
}

function saveCurrentOrder(container) {
  const ids = Array.from(container.querySelectorAll(".block-card")).map((card) =>
    String(card.dataset.blockId || "")
  );
  localStorage.setItem(ORDER_STORAGE_KEY, JSON.stringify(ids));
}

/* ---------------------- ПЕРЕТАСКИВАНИЕ СТОЛБЦОВ ---------------------- */

function initColumnDrag(table) {
  const headerRow = table.querySelector("thead tr");
  if (!headerRow) return;

  const headers = Array.from(headerRow.children);
  let srcIndex = null;

  headers.forEach((th) => {
    th.setAttribute("draggable", "true");

    th.addEventListener("dragstart", (event) => {
      event.stopPropagation();
      const currentHeaders = Array.from(headerRow.children);
      srcIndex = currentHeaders.indexOf(th);
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", "");
    });

    th.addEventListener("dragover", (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
    });

    th.addEventListener("drop", (event) => {
      event.preventDefault();
      event.stopPropagation();

      if (srcIndex === null) return;

      const currentHeaders = Array.from(headerRow.children);
      const targetIndex = currentHeaders.indexOf(th);
      if (targetIndex === -1 || targetIndex === srcIndex) {
        srcIndex = null;
        return;
      }

      moveTableColumn(table, srcIndex, targetIndex);
      srcIndex = null;
    });
  });
}

function moveTableColumn(table, fromIndex, toIndex) {
  const rows = table.querySelectorAll("tr");
  rows.forEach((row) => {
    const cells = Array.from(row.children);
    if (
      fromIndex < 0 ||
      fromIndex >= cells.length ||
      toIndex < 0 ||
      toIndex >= cells.length
    ) {
      return;
    }

    const fromCell = cells[fromIndex];
    const targetCell = cells[toIndex];

    if (fromIndex < toIndex) {
      row.insertBefore(fromCell, targetCell.nextSibling);
    } else {
      row.insertBefore(fromCell, targetCell);
    }
  });
}

/* ---------------------- СОРТИРОВКА ТАБЛИЦ ---------------------- */

function enableTableSorting(table, numericIndices) {
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  if (!thead || !tbody) return;

  const headers = Array.from(thead.querySelectorAll("th"));

  headers.forEach((th, colIndex) => {
    th.classList.add("sortable");

    th.addEventListener("click", () => {
      const currentOrder = th.dataset.sortOrder === "asc" ? "desc" : "asc";

      headers.forEach((h) => {
        h.removeAttribute("data-sort-order");
        h.classList.remove("sorted-asc", "sorted-desc");
      });

      th.dataset.sortOrder = currentOrder;
      th.classList.add(
        currentOrder === "asc" ? "sorted-asc" : "sorted-desc"
      );

      const rowsArray = Array.from(tbody.querySelectorAll("tr"));

      rowsArray.sort((rowA, rowB) => {
        const cellA = rowA.children[colIndex];
        const cellB = rowB.children[colIndex];

        const aText = cellA ? cellA.textContent.trim() : "";
        const bText = cellB ? cellB.textContent.trim() : "";

        const isNumeric =
          cellA && cellA.classList && cellA.classList.contains("cell-number");

        if (isNumeric) {
          const aVal = parseFloat(
            aText.replace(/\s/g, "").replace(",", ".")
          );
          const bVal = parseFloat(
            bText.replace(/\s/g, "").replace(",", ".")
          );

          const aNum = Number.isNaN(aVal) ? 0 : aVal;
          const bNum = Number.isNaN(bVal) ? 0 : bVal;

          return currentOrder === "asc" ? aNum - bNum : bNum - aNum;
        } else {
          return currentOrder === "asc"
            ? aText.localeCompare(bText)
            : bText.localeCompare(aText);
        }
      });

      rowsArray.forEach((row) => tbody.appendChild(row));
    });
  });
}

/* ---------------------- УТИЛИТЫ ---------------------- */

function formatNumberWithSpaces(value) {
  const str = String(value).trim();
  if (!str) return "";

  const num = parseFloat(str.replace(/\s/g, "").replace(",", "."));
  if (Number.isNaN(num)) return str;

  const [intPartRaw, fracPartRaw] = String(num).split(".");
  const intFormatted = intPartRaw.replace(/\B(?=(\d{3})+(?!\d))/g, " ");

  if (fracPartRaw !== undefined) {
    return `${intFormatted},${fracPartRaw}`;
  }
  return intFormatted;
}

function formatDateRu(dateStr) {
  // ожидаем YYYY-MM-DD
  if (!dateStr || typeof dateStr !== "string") return dateStr || "";
  const parts = dateStr.split("-");
  if (parts.length !== 3) return dateStr;
  const [y, m, d] = parts;
  return `${d}.${m}.${y}`;
}
