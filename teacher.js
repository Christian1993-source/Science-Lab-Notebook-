const TEACHER_EMAIL_SESSION_KEY = "libretaLaboratorio.teacherEmailSession";
const TEACHER_EMAIL_LOCAL_KEY = "libretaLaboratorio.teacherEmail";

const elements = {
  teacherEmail: document.getElementById("teacherEmail"),
  saveTeacherBtn: document.getElementById("saveTeacherBtn"),
  teacherSaveState: document.getElementById("teacherSaveState")
};

let restrictionAlertAt = 0;

init();

function init() {
  attachRestrictions();
  const savedEmail = readTeacherEmail();
  if (savedEmail) {
    elements.teacherEmail.value = savedEmail;
  }

  elements.saveTeacherBtn.addEventListener("click", onSaveEmail);
}

function readTeacherEmail() {
  return (
    String(sessionStorage.getItem(TEACHER_EMAIL_SESSION_KEY) || "").trim() ||
    String(localStorage.getItem(TEACHER_EMAIL_LOCAL_KEY) || "").trim()
  );
}

function onSaveEmail() {
  const email = String(elements.teacherEmail.value || "").trim();
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  if (!emailPattern.test(email)) {
    elements.teacherSaveState.textContent = "Please enter a valid email address.";
    return;
  }

  sessionStorage.setItem(TEACHER_EMAIL_SESSION_KEY, email);
  localStorage.setItem(TEACHER_EMAIL_LOCAL_KEY, email);
  elements.teacherSaveState.textContent = "Teacher email saved. Students can now submit final reports to this address.";
}

function attachRestrictions() {
  const blockEvent = (event) => {
    event.preventDefault();
    event.stopPropagation();
    showRestrictionAlert();
  };

  ["paste", "copy", "cut", "drop", "dragstart"].forEach((eventName) => {
    document.addEventListener(eventName, blockEvent, true);
  });

  document.addEventListener(
    "beforeinput",
    (event) => {
      const blockedTypes = new Set([
        "insertFromPaste",
        "insertFromDrop",
        "insertFromYank",
        "deleteByCut",
        "insertFromPasteAsQuotation"
      ]);
      if (blockedTypes.has(event.inputType)) {
        blockEvent(event);
      }
    },
    true
  );

  document.addEventListener(
    "keydown",
    (event) => {
      const key = event.key.toLowerCase();
      const withCommandKey = event.ctrlKey || event.metaKey;
      const blockedShortcuts = withCommandKey && ["c", "v", "x", "insert"].includes(key);
      const shiftInsert = event.shiftKey && key === "insert";
      if (blockedShortcuts || shiftInsert) {
        blockEvent(event);
      }
    },
    true
  );

  document.addEventListener("contextmenu", blockEvent, true);
  document.addEventListener("selectstart", blockEvent, true);
}

function showRestrictionAlert() {
  const now = Date.now();
  if (now - restrictionAlertAt < 1500) {
    return;
  }
  restrictionAlertAt = now;
  window.alert("Copy and paste are disabled. Please type the email manually.");
}
