const state = {
  workflows: [],
  statuses: {},
  selectedId: "hadith",
  quranOutputFiles: [],
  selectedQuranFiles: new Set(),
  duaOutputFiles: [],
  selectedDuaFiles: new Set(),
  pollTimer: null,
};

const workflowNav = document.getElementById("workflowNav");
const workflowTitle = document.getElementById("workflowTitle");
const workflowDescription = document.getElementById("workflowDescription");
const statusPill = document.getElementById("statusPill");
const dryRunRow = document.getElementById("dryRunRow");
const dryRunToggle = document.getElementById("dryRunToggle");
const runButton = document.getElementById("runButton");
const inputPaths = document.getElementById("inputPaths");
const outputPaths = document.getElementById("outputPaths");
const logOutput = document.getElementById("logOutput");
const exitCode = document.getElementById("exitCode");
const quranJsonPanel = document.getElementById("quranJsonPanel");
const quranUpdateStatus = document.getElementById("quranUpdateStatus");
const quranStructureSelect = document.getElementById("quranStructureSelect");
const quranOutputFiles = document.getElementById("quranOutputFiles");
const selectAllQuranFiles = document.getElementById("selectAllQuranFiles");
const updateQuranJsonButton = document.getElementById("updateQuranJsonButton");
const quranUpdateLog = document.getElementById("quranUpdateLog");
const duaJsonPanel = document.getElementById("duaJsonPanel");
const duaUpdateStatus = document.getElementById("duaUpdateStatus");
const duaOutputFiles = document.getElementById("duaOutputFiles");
const selectAllDuaFiles = document.getElementById("selectAllDuaFiles");
const updateDuaJsonButton = document.getElementById("updateDuaJsonButton");
const duaUpdateLog = document.getElementById("duaUpdateLog");

async function fetchJson(url, options) {
  const response = await fetch(url, options);
  if (!response.ok) {
    throw new Error(`Request failed: ${response.status}`);
  }
  return response.json();
}

async function loadWorkflows() {
  const payload = await fetchJson("/api/workflows");
  state.workflows = payload.workflows;
  state.statuses = payload.statuses;
  await loadQuranOutputFiles();
  await loadDuaOutputFiles();
  render();
  startPolling();
}

async function loadQuranOutputFiles() {
  const payload = await fetchJson("/api/quran-output-files");
  state.quranOutputFiles = payload.files || [];
  state.selectedQuranFiles = new Set(state.quranOutputFiles);
}

async function loadDuaOutputFiles() {
  const payload = await fetchJson("/api/dua-output-files");
  state.duaOutputFiles = payload.files || [];
  state.selectedDuaFiles = new Set(state.duaOutputFiles);
}

function selectedWorkflow() {
  return state.workflows.find((workflow) => workflow.id === state.selectedId) || state.workflows[0];
}

function selectedStatus() {
  return state.statuses[state.selectedId] || { state: "idle", logs: [] };
}

function render() {
  const workflow = selectedWorkflow();
  if (!workflow) return;
  const status = selectedStatus();
  const quranUpdate = state.statuses["quran-update-json"] || { state: "idle", logs: [] };
  const duaUpdate = state.statuses["dua-update-json"] || { state: "idle", logs: [] };

  workflowNav.innerHTML = state.workflows.map((item) => `
    <button class="nav-button ${item.id === workflow.id ? "active" : ""}" type="button" data-id="${item.id}">
      ${item.name}
    </button>
  `).join("");

  workflowTitle.textContent = workflow.name;
  workflowDescription.textContent = workflow.description;
  statusPill.textContent = labelForState(status.state);
  statusPill.className = `status-pill ${status.state}`;
  dryRunRow.style.display = workflow.supports_dry_run ? "flex" : "none";
  runButton.disabled = status.state === "running";
  runButton.textContent = status.state === "running" ? "Running..." : `Run ${workflow.name}`;
  inputPaths.innerHTML = workflow.input_paths.map(renderPathItem).join("");
  outputPaths.innerHTML = (status.outputs || workflow.output_paths.map((path) => ({ path, exists: false }))).map(renderPathItem).join("");
  logOutput.textContent = status.logs && status.logs.length ? status.logs.join("\n") : "No run started yet.";
  exitCode.textContent = status.exit_code === null || status.exit_code === undefined ? "" : `Exit code: ${status.exit_code}`;
  renderQuranJsonPanel(workflow, quranUpdate);
  renderDuaJsonPanel(workflow, duaUpdate);

  document.querySelectorAll(".nav-button").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedId = button.dataset.id;
      render();
    });
  });
}

function renderDuaJsonPanel(workflow, duaUpdate) {
  const visible = workflow.id === "dua";
  duaJsonPanel.classList.toggle("visible", visible);
  if (!visible) return;

  duaUpdateStatus.textContent = labelForState(duaUpdate.state);
  duaUpdateStatus.className = `status-pill ${duaUpdate.state}`;
  duaOutputFiles.innerHTML = state.duaOutputFiles.length
    ? state.duaOutputFiles.map(renderDuaFileOption).join("")
    : "<p>No .xlsx files found in DUA CONTENT\\OUTPUT. Run conversion first.</p>";

  document.querySelectorAll(".dua-file-checkbox").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        state.selectedDuaFiles.add(checkbox.value);
      } else {
        state.selectedDuaFiles.delete(checkbox.value);
      }
      render();
    });
  });

  const selectedCount = state.selectedDuaFiles.size;
  updateDuaJsonButton.disabled = duaUpdate.state === "running" || selectedCount === 0;
  updateDuaJsonButton.textContent = duaUpdate.state === "running"
    ? "Updating..."
    : `Update ${selectedCount} selected`;
  duaUpdateLog.textContent = duaUpdate.logs && duaUpdate.logs.length
    ? duaUpdate.logs.join("\n")
    : "No JSON update started yet.";
}

function renderQuranJsonPanel(workflow, quranUpdate) {
  const visible = workflow.id === "quran";
  quranJsonPanel.classList.toggle("visible", visible);
  if (!visible) return;

  quranUpdateStatus.textContent = labelForState(quranUpdate.state);
  quranUpdateStatus.className = `status-pill ${quranUpdate.state}`;
  quranOutputFiles.innerHTML = state.quranOutputFiles.length
    ? state.quranOutputFiles.map(renderQuranFileOption).join("")
    : "<p>No .xlsx files found in QURAN CONTENT\\OUTPUT.</p>";

  document.querySelectorAll(".quran-file-checkbox").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        state.selectedQuranFiles.add(checkbox.value);
      } else {
        state.selectedQuranFiles.delete(checkbox.value);
      }
      render();
    });
  });

  const selectedCount = state.selectedQuranFiles.size;
  updateQuranJsonButton.disabled = quranUpdate.state === "running" || selectedCount === 0;
  updateQuranJsonButton.textContent = quranUpdate.state === "running"
    ? "Updating..."
    : `Update ${selectedCount} selected`;
  quranUpdateLog.textContent = quranUpdate.logs && quranUpdate.logs.length
    ? quranUpdate.logs.join("\n")
    : "No JSON update started yet.";
}

function renderQuranFileOption(fileName) {
  const checked = state.selectedQuranFiles.has(fileName) ? "checked" : "";
  return `
    <label class="file-option">
      <input class="quran-file-checkbox" type="checkbox" value="${escapeHtml(fileName)}" ${checked}>
      <span>${escapeHtml(fileName)}</span>
    </label>
  `;
}

function renderDuaFileOption(fileName) {
  const checked = state.selectedDuaFiles.has(fileName) ? "checked" : "";
  return `
    <label class="file-option">
      <input class="dua-file-checkbox" type="checkbox" value="${escapeHtml(fileName)}" ${checked}>
      <span>${escapeHtml(fileName)}</span>
    </label>
  `;
}

function renderPathItem(item) {
  const path = typeof item === "string" ? item : item.path;
  const exists = typeof item === "string" ? null : item.exists;
  const stateClass = exists === null ? "" : exists ? "exists" : "missing";
  const className = `path-item ${stateClass}`.trim();
  return `<li class="${className}">${escapeHtml(path)}</li>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function labelForState(value) {
  return {
    idle: "Idle",
    running: "Running",
    succeeded: "Succeeded",
    failed: "Failed",
  }[value] || value;
}

async function startWorkflow() {
  const workflow = selectedWorkflow();
  const options = { dry_run: Boolean(dryRunToggle.checked) };
  const status = await fetchJson(`/api/start/${workflow.id}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(options),
  });
  state.statuses[workflow.id] = status;
  if (workflow.id === "dua") {
    setTimeout(async () => {
      await loadDuaOutputFiles();
      render();
    }, 1500);
  }
  render();
}

async function startQuranJsonUpdate() {
  const files = Array.from(state.selectedQuranFiles);
  const status = await fetchJson("/api/start/quran-update-json", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      structure: quranStructureSelect.value,
      files,
    }),
  });
  state.statuses["quran-update-json"] = status;
  render();
}

async function startDuaJsonUpdate() {
  const files = Array.from(state.selectedDuaFiles);
  const status = await fetchJson("/api/start/dua-update-json", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ files }),
  });
  state.statuses["dua-update-json"] = status;
  render();
}

function startPolling() {
  if (state.pollTimer) clearInterval(state.pollTimer);
  state.pollTimer = setInterval(refreshSelectedStatus, 1000);
}

async function refreshSelectedStatus() {
  if (!state.selectedId) return;
  const status = await fetchJson(`/api/status/${state.selectedId}`);
  state.statuses[state.selectedId] = status;
  if (state.selectedId === "quran") {
    state.statuses["quran-update-json"] = await fetchJson("/api/status/quran-update-json");
  }
  if (state.selectedId === "dua") {
    state.statuses["dua-update-json"] = await fetchJson("/api/status/dua-update-json");
    if (state.statuses.dua && state.statuses.dua.state !== "running") {
      await loadDuaOutputFiles();
    }
  }
  render();
}

runButton.addEventListener("click", () => {
  startWorkflow().catch((error) => {
    logOutput.textContent = error.message;
  });
});

selectAllQuranFiles.addEventListener("click", () => {
  const allSelected = state.selectedQuranFiles.size === state.quranOutputFiles.length;
  state.selectedQuranFiles = allSelected ? new Set() : new Set(state.quranOutputFiles);
  render();
});

selectAllDuaFiles.addEventListener("click", () => {
  const allSelected = state.selectedDuaFiles.size === state.duaOutputFiles.length;
  state.selectedDuaFiles = allSelected ? new Set() : new Set(state.duaOutputFiles);
  render();
});

updateQuranJsonButton.addEventListener("click", () => {
  startQuranJsonUpdate().catch((error) => {
    quranUpdateLog.textContent = error.message;
  });
});

updateDuaJsonButton.addEventListener("click", () => {
  startDuaJsonUpdate().catch((error) => {
    duaUpdateLog.textContent = error.message;
  });
});

loadWorkflows().catch((error) => {
  workflowTitle.textContent = "Dashboard failed to load";
  workflowDescription.textContent = error.message;
});
