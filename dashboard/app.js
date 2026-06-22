const state = {
  workflows: [],
  statuses: {},
  selectedId: "hadith",
  hadithReconciledFiles: [],
  selectedHadithReconciledFiles: new Set(),
  hadithBookWiseFinalFiles: [],
  selectedHadithBookWiseFinalFiles: new Set(),
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
const logOutput = document.getElementById("logOutput");
const exitCode = document.getElementById("exitCode");
const hadithUploadPanel = document.getElementById("hadithUploadPanel");
const hadithUploadInput = document.getElementById("hadithUploadInput");
const hadithUploadButton = document.getElementById("hadithUploadButton");
const hadithUploadLog = document.getElementById("hadithUploadLog");
const hadithFinalPanel = document.getElementById("hadithFinalPanel");
const hadithFinalStatus = document.getElementById("hadithFinalStatus");
const hadithReconciledFiles = document.getElementById("hadithReconciledFiles");
const selectAllHadithReconciledFiles = document.getElementById("selectAllHadithReconciledFiles");
const buildHadithFinalButton = document.getElementById("buildHadithFinalButton");
const hadithFinalLog = document.getElementById("hadithFinalLog");
const hadithUpdatedPanel = document.getElementById("hadithUpdatedPanel");
const hadithUpdatedStatus = document.getElementById("hadithUpdatedStatus");
const hadithBookWiseFinalFiles = document.getElementById("hadithBookWiseFinalFiles");
const selectAllHadithBookWiseFinalFiles = document.getElementById("selectAllHadithBookWiseFinalFiles");
const updateHadithContentButton = document.getElementById("updateHadithContentButton");
const hadithUpdatedLog = document.getElementById("hadithUpdatedLog");
const quranJsonPanel = document.getElementById("quranJsonPanel");
const quranUploadPanel = document.getElementById("quranUploadPanel");
const quranUploadInput = document.getElementById("quranUploadInput");
const quranUploadButton = document.getElementById("quranUploadButton");
const quranUploadLog = document.getElementById("quranUploadLog");
const quranUpdateStatus = document.getElementById("quranUpdateStatus");
const quranStructureSelect = document.getElementById("quranStructureSelect");
const quranOutputFiles = document.getElementById("quranOutputFiles");
const selectAllQuranFiles = document.getElementById("selectAllQuranFiles");
const updateQuranJsonButton = document.getElementById("updateQuranJsonButton");
const quranUpdateLog = document.getElementById("quranUpdateLog");
const duaJsonPanel = document.getElementById("duaJsonPanel");
const duaUploadPanel = document.getElementById("duaUploadPanel");
const duaUploadInput = document.getElementById("duaUploadInput");
const duaUploadButton = document.getElementById("duaUploadButton");
const duaUploadLog = document.getElementById("duaUploadLog");
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
  await loadHadithFiles();
  await loadQuranOutputFiles();
  await loadDuaOutputFiles();
  render();
  startPolling();
}

async function loadHadithFiles() {
  const previousReconciledFiles = state.hadithReconciledFiles;
  const previousReconciledSelection = state.selectedHadithReconciledFiles;
  const previousFinalFiles = state.hadithBookWiseFinalFiles;
  const previousFinalSelection = state.selectedHadithBookWiseFinalFiles;
  const reconciledPayload = await fetchJson("/api/hadith-reconciled-files");
  const finalPayload = await fetchJson("/api/hadith-book-wise-final-files");
  state.hadithReconciledFiles = reconciledPayload.files || [];
  state.selectedHadithReconciledFiles = nextSelection(
    state.hadithReconciledFiles,
    previousReconciledFiles,
    previousReconciledSelection,
  );
  state.hadithBookWiseFinalFiles = finalPayload.files || [];
  state.selectedHadithBookWiseFinalFiles = nextSelection(
    state.hadithBookWiseFinalFiles,
    previousFinalFiles,
    previousFinalSelection,
  );
}

function nextSelection(currentFiles, previousFiles, previousSelection) {
  const allWereSelected = previousFiles.length === 0 || previousFiles.length === previousSelection.size;
  if (allWereSelected) {
    return new Set(currentFiles);
  }
  return new Set(currentFiles.filter((fileName) => previousSelection.has(fileName)));
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
  const hadithFinal = state.statuses["hadith-book-wise-final"] || { state: "idle", logs: [] };
  const hadithUpdated = state.statuses["hadith-updated-content"] || { state: "idle", logs: [] };
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
  runButton.textContent = status.state === "running" ? "Workflow Running..." : `Run ${workflow.name}`;
  logOutput.textContent = status.logs && status.logs.length ? status.logs.join("\n") : "Workflow activity will appear here.";
  exitCode.textContent = status.exit_code === null || status.exit_code === undefined ? "" : `Exit code: ${status.exit_code}`;
  renderUploadPanels(workflow);
  renderHadithPanels(workflow, hadithFinal, hadithUpdated);
  renderQuranJsonPanel(workflow, quranUpdate);
  renderDuaJsonPanel(workflow, duaUpdate);

  document.querySelectorAll(".nav-button").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedId = button.dataset.id;
      render();
    });
  });
}

function renderUploadPanels(workflow) {
  hadithUploadPanel.classList.toggle("visible", workflow.id === "hadith");
  quranUploadPanel.classList.toggle("visible", workflow.id === "quran");
  duaUploadPanel.classList.toggle("visible", workflow.id === "dua");
}

function renderHadithPanels(workflow, hadithFinal, hadithUpdated) {
  const visible = workflow.id === "hadith";
  hadithFinalPanel.classList.toggle("visible", visible);
  hadithUpdatedPanel.classList.toggle("visible", visible);
  if (!visible) return;

  hadithFinalStatus.textContent = labelForState(hadithFinal.state);
  hadithFinalStatus.className = `status-pill ${hadithFinal.state}`;
  hadithReconciledFiles.innerHTML = state.hadithReconciledFiles.length
    ? state.hadithReconciledFiles.map((fileName) => renderFileOption(fileName, "hadith-reconciled-file-checkbox", state.selectedHadithReconciledFiles)).join("")
    : "<p>No .xlsx files found in SINGLE HADITH CONTENT\\reconciled_output.</p>";

  document.querySelectorAll(".hadith-reconciled-file-checkbox").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      updateSelection(state.selectedHadithReconciledFiles, checkbox);
      render();
    });
  });

  const finalCount = state.selectedHadithReconciledFiles.size;
  buildHadithFinalButton.disabled = hadithFinal.state === "running" || finalCount === 0;
  buildHadithFinalButton.textContent = hadithFinal.state === "running"
    ? "Building..."
    : `Build ${finalCount} selected`;
  hadithFinalLog.textContent = hadithFinal.logs && hadithFinal.logs.length
    ? hadithFinal.logs.join("\n")
    : "Book-wise final activity will appear here.";

  hadithUpdatedStatus.textContent = labelForState(hadithUpdated.state);
  hadithUpdatedStatus.className = `status-pill ${hadithUpdated.state}`;
  hadithBookWiseFinalFiles.innerHTML = state.hadithBookWiseFinalFiles.length
    ? state.hadithBookWiseFinalFiles.map((fileName) => renderFileOption(fileName, "hadith-final-file-checkbox", state.selectedHadithBookWiseFinalFiles)).join("")
    : "<p>No .xlsx files found in SINGLE HADITH CONTENT\\BOOK WISE FINAL.</p>";

  document.querySelectorAll(".hadith-final-file-checkbox").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      updateSelection(state.selectedHadithBookWiseFinalFiles, checkbox);
      render();
    });
  });

  const updatedCount = state.selectedHadithBookWiseFinalFiles.size;
  updateHadithContentButton.disabled = hadithUpdated.state === "running" || updatedCount === 0;
  updateHadithContentButton.textContent = hadithUpdated.state === "running"
    ? "Updating..."
    : `Update ${updatedCount} selected`;
  hadithUpdatedLog.textContent = hadithUpdated.logs && hadithUpdated.logs.length
    ? hadithUpdated.logs.join("\n")
    : "Final content update activity will appear here.";
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
    : "Structured output activity will appear here.";
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
    : "Structured output activity will appear here.";
}

function renderQuranFileOption(fileName) {
  return renderFileOption(fileName, "quran-file-checkbox", state.selectedQuranFiles);
}

function renderDuaFileOption(fileName) {
  return renderFileOption(fileName, "dua-file-checkbox", state.selectedDuaFiles);
}

function renderFileOption(fileName, className, selectedSet) {
  const checked = selectedSet.has(fileName) ? "checked" : "";
  return `
    <label class="file-option">
      <input class="${className}" type="checkbox" value="${escapeHtml(fileName)}" ${checked}>
      <span>${escapeHtml(fileName)}</span>
    </label>
  `;
}

function updateSelection(selectedSet, checkbox) {
  if (checkbox.checked) {
    selectedSet.add(checkbox.value);
  } else {
    selectedSet.delete(checkbox.value);
  }
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
  if (workflow.id === "hadith") {
    setTimeout(async () => {
      await loadHadithFiles();
      render();
    }, 1500);
  }
  render();
}

async function startHadithBookWiseFinal() {
  const files = Array.from(state.selectedHadithReconciledFiles);
  const status = await fetchJson("/api/start/hadith-book-wise-final", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ files }),
  });
  state.statuses["hadith-book-wise-final"] = status;
  render();
}

async function startHadithUpdatedContent() {
  const files = Array.from(state.selectedHadithBookWiseFinalFiles);
  const status = await fetchJson("/api/start/hadith-updated-content", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ files }),
  });
  state.statuses["hadith-updated-content"] = status;
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

async function uploadWorkflowFiles(workflowId, inputElement, logElement) {
  let files = Array.from(inputElement.files || []);
  if (!files.length) {
    logElement.textContent = "Choose one or more Excel workbooks first.";
    return;
  }

  const targetPayload = await fetchJson(`/api/upload-target-files/${workflowId}`);
  const existingFiles = new Set(targetPayload.files || []);
  const replacing = files.filter((file) => existingFiles.has(file.name));
  let overwrite = false;
  if (replacing.length) {
    overwrite = window.confirm(`Replace ${replacing.length} existing file(s)?`);
    if (!overwrite) {
      files = files.filter((file) => !existingFiles.has(file.name));
    }
  }

  if (!files.length) {
    logElement.textContent = "No files uploaded.";
    return;
  }

  const encodedFiles = [];
  for (const file of files) {
    encodedFiles.push({
      name: file.name,
      content: await readFileAsDataUrl(file),
    });
  }

  const result = await fetchJson(`/api/upload/${workflowId}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ overwrite, files: encodedFiles }),
  });

  inputElement.value = "";
  logElement.textContent = uploadSummary(result);
  await refreshFileLists();
  render();
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(reader.result));
    reader.addEventListener("error", () => reject(reader.error || new Error("File read failed.")));
    reader.readAsDataURL(file);
  });
}

function uploadSummary(result) {
  const lines = [];
  const saved = result.saved || [];
  const skipped = result.skipped || [];
  lines.push(saved.length ? `Uploaded: ${saved.join(", ")}` : "Uploaded: none");
  if (skipped.length) {
    lines.push("Skipped:");
    skipped.forEach((item) => lines.push(`${item.name}: ${item.reason}`));
  }
  return lines.join("\n");
}

async function refreshFileLists() {
  await loadHadithFiles();
  await loadQuranOutputFiles();
  await loadDuaOutputFiles();
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
  if (state.selectedId === "hadith") {
    state.statuses["hadith-book-wise-final"] = await fetchJson("/api/status/hadith-book-wise-final");
    state.statuses["hadith-updated-content"] = await fetchJson("/api/status/hadith-updated-content");
    if (state.statuses["hadith-book-wise-final"].state !== "running") {
      await loadHadithFiles();
    }
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

selectAllHadithReconciledFiles.addEventListener("click", () => {
  const allSelected = state.selectedHadithReconciledFiles.size === state.hadithReconciledFiles.length;
  state.selectedHadithReconciledFiles = allSelected ? new Set() : new Set(state.hadithReconciledFiles);
  render();
});

selectAllHadithBookWiseFinalFiles.addEventListener("click", () => {
  const allSelected = state.selectedHadithBookWiseFinalFiles.size === state.hadithBookWiseFinalFiles.length;
  state.selectedHadithBookWiseFinalFiles = allSelected ? new Set() : new Set(state.hadithBookWiseFinalFiles);
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

buildHadithFinalButton.addEventListener("click", () => {
  startHadithBookWiseFinal().catch((error) => {
    hadithFinalLog.textContent = error.message;
  });
});

updateHadithContentButton.addEventListener("click", () => {
  startHadithUpdatedContent().catch((error) => {
    hadithUpdatedLog.textContent = error.message;
  });
});

hadithUploadButton.addEventListener("click", () => {
  uploadWorkflowFiles("hadith", hadithUploadInput, hadithUploadLog).catch((error) => {
    hadithUploadLog.textContent = error.message;
  });
});

quranUploadButton.addEventListener("click", () => {
  uploadWorkflowFiles("quran", quranUploadInput, quranUploadLog).catch((error) => {
    quranUploadLog.textContent = error.message;
  });
});

duaUploadButton.addEventListener("click", () => {
  uploadWorkflowFiles("dua", duaUploadInput, duaUploadLog).catch((error) => {
    duaUploadLog.textContent = error.message;
  });
});

loadWorkflows().catch((error) => {
  workflowTitle.textContent = "Dashboard failed to load";
  workflowDescription.textContent = error.message;
});
