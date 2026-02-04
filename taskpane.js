/* global Office, Word, Excel */
const OLLAMA_BASE_URL = "/ollama";

const statusEl = document.getElementById("status");
const modelEl = document.getElementById("model");
const promptEl = document.getElementById("prompt");
const outputEl = document.getElementById("output");
const refreshEl = document.getElementById("refresh");
const generateEl = document.getElementById("generate");
const insertEl = document.getElementById("insert");

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
        range.insertText(text, Word.InsertLocation.replace);
        await context.sync();
      });
      setStatus("Insertado en Word");
    } else if (host === Office.HostType.Excel) {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = [[text]];
        await context.sync();
      });
      setStatus("Insertado en Excel");
    } else {
      setStatus("Host no compatible", true);
    }
  } catch (err) {
    setStatus(`Error al insertar: ${err.message}`, true);
  }
}

Office.onReady(() => {
  refreshEl.addEventListener("click", fetchModels);
  generateEl.addEventListener("click", generate);
  insertEl.addEventListener("click", insertIntoDocument);
  fetchModels();
});
