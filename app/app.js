const form = document.getElementById("cvForm");
const template = document.getElementById("template");
const resumes = document.getElementById("resumes");
const templateLabel = document.getElementById("templateLabel");
const resumeLabel = document.getElementById("resumeLabel");
const statusBox = document.getElementById("status");
const resultList = document.getElementById("resultList");
const submitButton = document.getElementById("submitButton");
const buttonLabel = submitButton.querySelector(".button-label");
const nodes = [...document.querySelectorAll(".node")];

const defaultButtonText = buttonLabel.textContent;

template.addEventListener("change", () => {
  templateLabel.textContent = template.files[0]?.name || "Choose DOCX template";
});

resumes.addEventListener("change", () => {
  const count = resumes.files.length;
  resumeLabel.textContent = count ? `${count} resume${count > 1 ? "s" : ""} selected` : "Choose PDF or DOCX resumes";
});

function setStep(index) {
  nodes.forEach((node, i) => node.classList.toggle("active", i <= index));
}

function setStatus(message, state = "idle") {
  statusBox.className = `status ${state}`;
  statusBox.innerHTML = message;
}

function setProcessing(isProcessing) {
  form.classList.toggle("is-processing", isProcessing);
  submitButton.disabled = isProcessing;
  template.disabled = isProcessing;
  resumes.disabled = isProcessing;
  buttonLabel.textContent = isProcessing ? "Generating" : defaultButtonText;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderResults(payload) {
  setStatus(`Done. <a href="${escapeHtml(payload.zip)}">Download ZIP</a>`, "success");
  resultList.innerHTML = "";

  payload.results.forEach((item) => {
    const card = document.createElement("article");
    card.className = "result-card";

    const name = escapeHtml(item.name || item.source || "Generated CV");
    const position = escapeHtml(item.position || "Position not detected");
    const mappingLabel = item.ai_used ? "Gemini mapped" : "Fallback mapped";
    const reviewLabel = item.needs_review ? "Needs review" : "Ready";
    const mappingClass = item.ai_used ? "success" : "warning";
    const reviewClass = item.needs_review ? "warning" : "success";
    const aiError = item.ai_error ? `<small class="error-detail">${escapeHtml(item.ai_error)}</small>` : "";
    const extraction = item.extraction ? `<a href="${escapeHtml(item.extraction)}">Extraction JSON</a>` : "";

    card.innerHTML = `
      <div>
        <strong>${name}</strong>
        <div class="result-meta">${position}</div>
        <span class="pill ${mappingClass}">${mappingLabel}</span>
        <span class="pill ${reviewClass}">${reviewLabel}</span>
        ${aiError}
      </div>
      <div class="result-links">
        ${extraction}
        <a href="${escapeHtml(item.download)}">Download DOCX</a>
      </div>
    `;
    resultList.appendChild(card);
  });
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  resultList.innerHTML = "";
  setStep(1);
  setProcessing(true);
  setStatus("Uploading files", "processing");

  const body = new FormData();
  body.append("template", template.files[0]);
  [...resumes.files].forEach((file) => body.append("resumes", file));

  const timers = [
    setTimeout(() => {
      setStep(2);
      setStatus("Extracting resume text", "processing");
    }, 500),
    setTimeout(() => {
      setStep(3);
      setStatus("Mapping fields with Gemini", "processing");
    }, 1600),
    setTimeout(() => {
      setStep(4);
      setStatus("Building Word files", "processing");
    }, 2800),
  ];

  try {
    const response = await fetch("/api/process", { method: "POST", body });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Processing failed");
    setStep(4);
    renderResults(payload);
  } catch (error) {
    setStep(0);
    setStatus(escapeHtml(error.message), "error");
  } finally {
    timers.forEach(clearTimeout);
    setProcessing(false);
  }
});
