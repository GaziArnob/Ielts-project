const root = document.documentElement;
const themeToggle = document.getElementById("themeToggle");
const storageKey = "bandforge-theme";

function applyTheme(theme) {
  if (theme === "dark") {
    root.setAttribute("data-theme", "dark");
    if (themeToggle) themeToggle.textContent = "Light Mode";
    return;
  }
  root.removeAttribute("data-theme");
  if (themeToggle) themeToggle.textContent = "Dark Mode";
}

applyTheme(localStorage.getItem(storageKey) || "light");

if (themeToggle) {
  themeToggle.addEventListener("click", () => {
    const next = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
    localStorage.setItem(storageKey, next);
    applyTheme(next);
  });
}

function speakLines(lines) {
  if (!("speechSynthesis" in window)) return;
  window.speechSynthesis.cancel();
  lines.forEach((line, index) => {
    const utterance = new SpeechSynthesisUtterance(line);
    utterance.rate = 0.92;
    utterance.pitch = 1;
    utterance.volume = 1;
    utterance.lang = "en-US";
    utterance.onstart = () => {
      if (index === 0) document.body.classList.add("speaking-now");
    };
    utterance.onend = () => {
      if (index === lines.length - 1) document.body.classList.remove("speaking-now");
    };
    window.speechSynthesis.speak(utterance);
  });
}

const playListening = document.getElementById("playListening");
const stopAudio = document.getElementById("stopAudio");

if (playListening) {
  playListening.addEventListener("click", () => speakLines(window.__LISTENING_SCRIPT__ || []));
}

if (stopAudio) {
  stopAudio.addEventListener("click", () => {
    if ("speechSynthesis" in window) window.speechSynthesis.cancel();
  });
}

document.querySelectorAll(".prompt-play").forEach((button) => {
  button.addEventListener("click", () => {
    const text = button.dataset.text || "";
    if (text) speakLines([text]);
  });
});

const recordedBlobs = {};
let activeRecorder = null;
let activeChunks = [];
let activeField = "";

document.querySelectorAll(".record-strip").forEach((strip) => {
  const startBtn = strip.querySelector(".record-start");
  const stopBtn = strip.querySelector(".record-stop");
  const preview = document.getElementById(strip.dataset.preview);
  const fieldName = strip.dataset.field;

  if (!navigator.mediaDevices || !window.MediaRecorder) {
    startBtn.disabled = true;
    stopBtn.disabled = true;
    return;
  }

  startBtn.addEventListener("click", async () => {
    if (activeRecorder && activeRecorder.state === "recording") return;
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    activeField = fieldName;
    activeChunks = [];
    activeRecorder = new MediaRecorder(stream);
    activeRecorder.ondataavailable = (event) => {
      if (event.data && event.data.size > 0) activeChunks.push(event.data);
    };
    activeRecorder.onstop = () => {
      const blob = new Blob(activeChunks, { type: activeRecorder.mimeType || "audio/webm" });
      recordedBlobs[activeField] = blob;
      preview.src = URL.createObjectURL(blob);
      stream.getTracks().forEach((track) => track.stop());
    };
    activeRecorder.start();
  });

  stopBtn.addEventListener("click", () => {
    if (activeRecorder && activeRecorder.state === "recording") activeRecorder.stop();
  });
});

const speakingForm = document.getElementById("speakingForm");
if (speakingForm) {
  speakingForm.addEventListener("submit", async (event) => {
    if (!Object.keys(recordedBlobs).length) return;
    event.preventDefault();
    const formData = new FormData(speakingForm);
    Object.entries(recordedBlobs).forEach(([field, blob]) => {
      formData.set(field, blob, `${field}.webm`);
    });
    const response = await fetch(speakingForm.action, { method: "POST", body: formData });
    if (response.redirected) {
      window.location.href = response.url;
      return;
    }
    document.open();
    document.write(await response.text());
    document.close();
  });
}
