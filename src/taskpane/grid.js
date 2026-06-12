// Excel-style datasheet grid for the taskpane (think-cell's datasheet,
// simplified): editable cells, paste-from-Excel that expands the grid,
// auto-growing rows/columns, and automatic continuation of year sequences
// in the category header row.

export function createGrid(container, { rows = 8, cols = 6, onChange } = {}) {
  let data = emptyData(rows, cols);
  let table;

  function emptyData(r, c) {
    return Array.from({ length: r }, () => Array.from({ length: c }, () => ""));
  }

  function render() {
    container.innerHTML = "";
    table = document.createElement("table");
    table.className = "datasheet";
    data.forEach((row, r) => {
      const tr = document.createElement("tr");
      row.forEach((val, c) => {
        const td = document.createElement("td");
        if (r === 0 || c === 0) td.className = "head";
        const input = document.createElement("input");
        input.value = val;
        input.dataset.r = r;
        input.dataset.c = c;
        input.addEventListener("input", onInput);
        input.addEventListener("paste", onPaste);
        input.addEventListener("keydown", onKey);
        td.appendChild(input);
        tr.appendChild(td);
      });
      table.appendChild(tr);
    });
    container.appendChild(table);
  }

  function onInput(e) {
    const r = +e.target.dataset.r;
    const c = +e.target.dataset.c;
    data[r][c] = e.target.value;
    autofillHeader(r, c);
    growIfEdge(r, c);
    onChange?.();
  }

  // think-cell: when you enter data under an empty category header and the
  // previous headers form a numeric sequence (e.g. years), continue it.
  function autofillHeader(r, c) {
    if (r === 0 || c < 3 || data[0][c] !== "" || data[r][c] === "") return;
    const a = Number(data[0][c - 2]);
    const b = Number(data[0][c - 1]);
    if (data[0][c - 2] !== "" && data[0][c - 1] !== "" &&
        Number.isFinite(a) && Number.isFinite(b) && b !== a) {
      data[0][c] = String(b + (b - a));
      const input = cellInput(0, c);
      if (input) input.value = data[0][c];
    }
  }

  // Typing in the last row/column grows the sheet (think-cell's datasheet
  // grows with the used area).
  function growIfEdge(r, c) {
    let grew = false;
    if (r === data.length - 1 && data[r][c] !== "") {
      data.push(Array.from({ length: data[0].length }, () => ""));
      grew = true;
    }
    if (c === data[0].length - 1 && data[r][c] !== "") {
      data.forEach((row) => row.push(""));
      grew = true;
    }
    if (grew) {
      const focus = { r, c };
      render();
      cellInput(focus.r, focus.c)?.focus();
    }
  }

  function onPaste(e) {
    const text = (e.clipboardData || window.clipboardData).getData("text");
    if (!text || (!text.includes("\t") && !text.includes("\n"))) return;
    e.preventDefault();
    const r0 = +e.target.dataset.r;
    const c0 = +e.target.dataset.c;
    const rows = text.replace(/\r\n?/g, "\n").split("\n").filter((l) => l !== "");
    rows.forEach((line, dr) => {
      line.split("\t").forEach((cell, dc) => {
        const r = r0 + dr;
        const c = c0 + dc;
        while (data.length <= r + 1) data.push(Array.from({ length: data[0].length }, () => ""));
        while (data[0].length <= c + 1) data.forEach((row) => row.push(""));
        data[r][c] = cell.trim();
      });
    });
    render();
    onChange?.();
  }

  function onKey(e) {
    const r = +e.target.dataset.r;
    const c = +e.target.dataset.c;
    let target = null;
    if (e.key === "Enter" || (e.key === "ArrowDown" && !e.shiftKey)) target = [r + 1, c];
    else if (e.key === "ArrowUp") target = [r - 1, c];
    else if (e.key === "Tab") return; // native tab order is already row-major
    if (target) {
      const input = cellInput(target[0], target[1]);
      if (input) {
        e.preventDefault();
        input.focus();
        input.select();
      }
    }
  }

  function cellInput(r, c) {
    return container.querySelector(`input[data-r="${r}"][data-c="${c}"]`);
  }

  // Returns the used area as a 2D string array (trailing empty rows/cols
  // trimmed), mirroring how think-cell sizes the chart to the used range.
  function getData() {
    let maxR = -1, maxC = -1;
    data.forEach((row, r) =>
      row.forEach((cell, c) => {
        if (cell !== "") {
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        }
      })
    );
    if (maxR < 0) return [];
    return data.slice(0, maxR + 1).map((row) => row.slice(0, maxC + 1));
  }

  function setData(arr) {
    const r = Math.max(arr.length + 2, 8);
    const c = Math.max(...arr.map((row) => row.length), 4) + 2;
    data = emptyData(r, c);
    arr.forEach((row, ri) => row.forEach((cell, ci) => { data[ri][ci] = String(cell ?? ""); }));
    render();
    onChange?.();
  }

  function clear() {
    data = emptyData(rows, cols);
    render();
    onChange?.();
  }

  render();
  return { getData, setData, clear };
}
