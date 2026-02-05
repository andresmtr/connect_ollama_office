/* global Office, Word, Excel */
const OLLAMA_BASE_URL = "/ollama";

const statusEl = document.getElementById("status");
const modelEl = document.getElementById("model");
const promptEl = document.getElementById("prompt");
const outputEl = document.getElementById("output");
const refreshEl = document.getElementById("refresh");
const generateEl = document.getElementById("generate");
const insertEl = document.getElementById("insert");
const wordActionsEl = document.getElementById("word-actions");
const wordActionEl = document.getElementById("word-action");
const wordToneRowEl = document.getElementById("word-tone-row");
const wordToneEl = document.getElementById("word-tone");
const wordExtraEl = document.getElementById("word-extra");
const wordRunEl = document.getElementById("word-run");
const excelActionsEl = document.getElementById("excel-actions");
const excelActionEl = document.getElementById("excel-action");
const excelFormulaRowEl = document.getElementById("excel-formula-row");
const excelFormulaDescEl = document.getElementById("excel-formula-desc");
const excelExtraEl = document.getElementById("excel-extra");
const excelRunEl = document.getElementById("excel-run");
const excelOutputEl = document.getElementById("excel-output");
const excelParallelEl = document.getElementById("excel-parallel");

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.classList.toggle("error", isError);
}

async function fetchModels() {
  setStatus("Cargando modelos...");
  try {
    const response = await fetch(`${OLLAMA_BASE_URL}/api/tags`);
    if (!response.ok) {
      throw new Error(`Error HTTP ${response.status}`);
    }
    const data = await response.json();
    const models = (data.models || []).map((m) => m.name);

    modelEl.innerHTML = "";
    if (models.length === 0) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "No hay modelos";
      modelEl.appendChild(option);
    } else {
      models.forEach((name) => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        modelEl.appendChild(option);
      });
    }
    setStatus("Modelos listos");
  } catch (err) {
    setStatus(`No se pudo cargar modelos: ${err.message}`, true);
  }
}

async function generate() {
  const model = modelEl.value;
  const prompt = promptEl.value.trim();

  if (!model) {
    setStatus("Selecciona un modelo", true);
    return;
  }
  if (!prompt) {
    setStatus("Escribe un prompt", true);
    return;
  }

  setStatus("Generando...");
  outputEl.value = "";

  try {
    const response = await fetch(`${OLLAMA_BASE_URL}/api/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model,
        prompt,
        stream: false
      })
    });

    if (!response.ok) {
      throw new Error(`Error HTTP ${response.status}`);
    }

    const data = await response.json();
    outputEl.value = data.response || "";
    setStatus("Respuesta lista");
  } catch (err) {
    setStatus(`Error al generar: ${err.message}`, true);
  }
}

async function insertIntoDocument() {
  const text = outputEl.value.trim();
  if (!text) {
    setStatus("No hay respuesta para insertar", true);
    return;
  }

  const host = Office.context.host;
  try {
    if (host === Office.HostType.Word) {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(`\n${text}`, Word.InsertLocation.after);
        await context.sync();
      });
      setStatus("Insertado en Word (debajo)");
    } else if (host === Office.HostType.Excel) {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        const target = range.getOffsetRange(0, range.columnCount);
        target.values = [[text]];
        await context.sync();
      });
      setStatus("Insertado en Excel (celda derecha)");
    } else {
      setStatus("Host no compatible", true);
    }
  } catch (err) {
    setStatus(`Error al insertar: ${err.message}`, true);
  }
}

async function callOllama(model, prompt) {
  const response = await fetch(`${OLLAMA_BASE_URL}/api/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model,
      prompt,
      stream: false
    })
  });

  if (!response.ok) {
    throw new Error(`Error HTTP ${response.status}`);
  }

  const data = await response.json();
  return (data.response || "").trim();
}

function ensureModel() {
  const model = modelEl.value;
  if (!model) {
    setStatus("Selecciona un modelo", true);
    return null;
  }
  return model;
}

async function getWordSelectionText() {
  let text = "";
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    text = range.text || "";
  });
  return text.trim();
}

async function runWordAction() {
  const model = ensureModel();
  if (!model) return;

  const selection = await getWordSelectionText();
  if (!selection) {
    setStatus("Selecciona texto en Word", true);
    return;
  }

  const action = wordActionEl.value;
  const extra = wordExtraEl.value.trim();
  const tone = wordToneEl.value;

  let instruction = "";
  if (action === "correct") {
    instruction = "Corrige ortografía, gramática y estilo sin cambiar el significado.";
  } else if (action === "rewrite") {
    instruction = `Reescribe el texto con tono ${tone}.`;
  } else if (action === "summarize") {
    instruction = "Resume el texto en 3 a 5 líneas.";
  } else if (action === "keypoints") {
    instruction = "Extrae las ideas clave en viñetas claras.";
  }

  const prompt = [
    instruction,
    extra ? `Instrucciones extra: ${extra}` : "",
    "Texto:",
    selection
  ]
    .filter(Boolean)
    .join("\n\n");

  setStatus("Generando en Word...");
  try {
    const result = await callOllama(model, prompt);
    if (!result) {
      setStatus("Respuesta vacía", true);
      return;
    }

    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(`\n${result}`, Word.InsertLocation.after);
      await context.sync();
    });
    setStatus("Acción aplicada en Word");
  } catch (err) {
    setStatus(`Error en Word: ${err.message}`, true);
  }
}

function rangeToTSV(values) {
  return values
    .map((row) =>
      row
        .map((cell) => {
          if (cell === null || cell === undefined) return "";
          return String(cell).replace(/\t/g, " ").replace(/\n/g, " ");
        })
        .join("\t")
    )
    .join("\n");
}

function rowToText(row) {
  return row
    .map((cell) => {
      if (cell === null || cell === undefined) return "";
      return String(cell).trim();
    })
    .filter(Boolean)
    .join(" | ");
}

async function runPerRowAction({
  model,
  range,
  values,
  buildPrompt,
  onResult,
  concurrency,
  placeValues
}) {
  const output = Array(values.length).fill(null);
  const total = values.length;
  const limit = Math.max(1, Number(concurrency || 1));

  let index = 0;
  const workers = Array.from({ length: limit }).map(async () => {
    while (index < total) {
      const current = index;
      index += 1;
      const row = values[current] || [];
      const rowText = rowToText(row);
      if (!rowText) {
        output[current] = [""];
        continue;
      }

      const prompt = buildPrompt(rowText, current);
      const result = await callOllama(model, prompt);
      output[current] = [result || ""];
      if (onResult) {
        onResult(result, current);
      }
    }
  });

  await Promise.all(workers);
  await placeValues(output);
}

async function runExcelAction() {
  const model = ensureModel();
  if (!model) return;

  const action = excelActionEl.value;
  const extra = excelExtraEl.value.trim();

  setStatus("Leyendo selección de Excel...");
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load([
        "values",
        "address",
        "rowCount",
        "columnCount",
        "rowIndex",
        "columnIndex"
      ]);
      await context.sync();

      const values = range.values || [];
      const tsv = rangeToTSV(values);
      const outputMode = excelOutputEl.value;
      const concurrency = Number(excelParallelEl.value || 1);
      let prompt = "";

      if (action === "formula") {
        const desc = excelFormulaDescEl.value.trim();
        if (!desc) {
          throw new Error("Describe la fórmula en el campo de descripción.");
        }
        prompt = [
          "Crea una fórmula de Excel en inglés.",
          "Responde SOLO con la fórmula, sin explicaciones ni backticks.",
          `Descripción: ${desc}`,
          `Rango seleccionado: ${range.address}`,
          "Datos (TSV):",
          tsv
        ].join("\n");
      } else if (action === "summarize_range") {
        prompt = [
          "Resume el rango seleccionado en 1 a 3 frases.",
          extra ? `Instrucciones extra: ${extra}` : "",
          `Rango seleccionado: ${range.address}`,
          "Datos (TSV):",
          tsv
        ]
          .filter(Boolean)
          .join("\n");
      } else if (action === "stats") {
        prompt = [
          "Calcula min, max, promedio y conteo de los valores numéricos.",
          "Responde SOLO en JSON con claves: min, max, avg, count.",
          `Rango seleccionado: ${range.address}`,
          "Datos (TSV):",
          tsv
        ].join("\n");
      }

      setStatus("Generando en Excel...");

      const perRowActions = new Set([
        "summarize_row",
        "extract_numbers",
        "rewrite_row",
        "translate_row",
        "classify_row",
        "clean_row"
      ]);

      const normalizeOutput = (output) => {
        if (!Array.isArray(output)) {
          return [[output ?? ""]];
        }
        if (output.length === 0) {
          return [[""]];
        }
        if (!Array.isArray(output[0])) {
          return output.map((item) => [item ?? ""]);
        }
        const maxCols = Math.max(
          1,
          ...output.map((row) => (Array.isArray(row) ? row.length : 0))
        );
        return output.map((row) => {
          const safeRow = Array.isArray(row) ? row : [row ?? ""];
          if (safeRow.length === maxCols) return safeRow;
          const filled = safeRow.slice();
          while (filled.length < maxCols) filled.push("");
          return filled;
        });
      };

      const placeValues = async (output) => {
        const normalized = normalizeOutput(output);
        const rowCount = normalized.length || 1;
        const colCount = normalized[0].length || 1;
        if (outputMode === "new_sheet") {
          const sheet = context.workbook.worksheets.add(
            `Ollama_${Date.now()}`
          );
          const target = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
          target.values = normalized;
          sheet.activate();
          return;
        }

        const sheet = range.worksheet;
        let startRow = range.rowIndex;
        let startCol = range.columnIndex;
        if (outputMode === "right") {
          startCol = range.columnIndex + range.columnCount;
        } else if (outputMode === "below") {
          startRow = range.rowIndex + range.rowCount;
        }

        const targetRange = sheet.getRangeByIndexes(
          startRow,
          startCol,
          rowCount,
          colCount
        );
        targetRange.values = normalized;
      };

      if (perRowActions.has(action)) {
        await runPerRowAction({
          model,
          range,
          values,
          concurrency,
          placeValues,
          buildPrompt: (rowText) => {
            if (action === "summarize_row") {
              return [
                "Resume el texto en 1 frase.",
                extra ? `Instrucciones extra: ${extra}` : "",
                "Texto:",
                rowText
              ]
                .filter(Boolean)
                .join("\n");
            }
            if (action === "extract_numbers") {
              return [
                "Extrae y lista solo los valores numéricos relevantes.",
                "Si no hay números, responde vacío.",
                extra ? `Instrucciones extra: ${extra}` : "",
                "Texto:",
                rowText
              ]
                .filter(Boolean)
                .join("\n");
            }
            if (action === "rewrite_row") {
              return [
                "Reescribe el texto manteniendo el significado.",
                extra ? `Instrucciones extra: ${extra}` : "",
                "Texto:",
                rowText
              ]
                .filter(Boolean)
                .join("\n");
            }
            if (action === "translate_row") {
              return [
                "Traduce el texto al idioma indicado.",
                extra
                  ? `Idioma destino y notas: ${extra}`
                  : "Idioma destino: inglés.",
                "Texto:",
                rowText
              ].join("\n");
            }
            if (action === "classify_row") {
              return [
                "Clasifica el texto en una etiqueta corta.",
                extra
                  ? `Etiquetas posibles o reglas: ${extra}`
                  : "Si no hay etiquetas, inventa una categoría breve.",
                "Texto:",
                rowText
              ].join("\n");
            }
            return [
              "Normaliza el texto: quita espacios extra, corrige mayúsculas y elimina caracteres raros.",
              extra ? `Instrucciones extra: ${extra}` : "",
              "Texto:",
              rowText
            ]
              .filter(Boolean)
              .join("\n");
          }
        });
      } else if (action === "stats") {
        const result = await callOllama(model, prompt);
        if (!result) {
          throw new Error("Respuesta vacía");
        }
        let stats = null;
        try {
          stats = JSON.parse(result);
        } catch (err) {
          stats = null;
        }

        const output = stats && typeof stats === "object"
          ? [
              ["min", stats.min ?? ""],
              ["max", stats.max ?? ""],
              ["avg", stats.avg ?? ""],
              ["count", stats.count ?? ""]
            ]
          : [[result]];
        await placeValues(output);
      } else {
        const result = await callOllama(model, prompt);
        if (!result) {
          throw new Error("Respuesta vacía");
        }
        await placeValues([[result]]);
      }

      await context.sync();
    });

    setStatus("Acción aplicada en Excel");
  } catch (err) {
    setStatus(`Error en Excel: ${err.message}`, true);
  }
}

function toggleActionUI() {
  const host = Office.context.host;
  wordActionsEl.style.display = host === Office.HostType.Word ? "grid" : "none";
  excelActionsEl.style.display = host === Office.HostType.Excel ? "grid" : "none";
  wordToneRowEl.style.display =
    wordActionEl.value === "rewrite" ? "flex" : "none";
  excelFormulaRowEl.style.display =
    excelActionEl.value === "formula" ? "block" : "none";
}

Office.onReady(() => {
  refreshEl.addEventListener("click", fetchModels);
  generateEl.addEventListener("click", generate);
  insertEl.addEventListener("click", insertIntoDocument);
  wordRunEl.addEventListener("click", runWordAction);
  excelRunEl.addEventListener("click", runExcelAction);
  wordActionEl.addEventListener("change", toggleActionUI);
  excelActionEl.addEventListener("change", toggleActionUI);
  fetchModels();
  toggleActionUI();
});
