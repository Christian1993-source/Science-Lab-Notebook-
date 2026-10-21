/* Quiz content lives in practice-data.json. Scores are calculated on this device. */
"use strict";

function gradeQuiz(quiz, answers) {
  const results = quiz.questions.map((question, index) => {
    const answer = answers[index];
    const answered = answer !== undefined && answer !== null && String(answer).trim() !== "";
    const correct = answered && (question.type === "numeric"
      ? Number.isFinite(Number(answer)) && Math.abs(Number(answer) - question.answer) <= (question.tolerance || 0)
      : Number(answer) === question.answer);
    return { correct, answered, answer };
  });
  const score = results.filter((result) => result.correct).length;
  return { score, total: results.length, percent: results.length ? Math.round(score / results.length * 100) : 0, results };
}

function validatePracticeData(data) {
  if (!Array.isArray(data.courses)) throw new Error("Missing courses");
  const courseIds = new Set();
  data.courses.forEach((course) => {
    if (!course.id || courseIds.has(course.id) || !course.title || !Array.isArray(course.quizzes)) throw new Error("Invalid course");
    courseIds.add(course.id);
    const quizIds = new Set();
    course.quizzes.forEach((quiz) => {
      if (!quiz.id || quizIds.has(quiz.id) || !quiz.title || !Array.isArray(quiz.questions) || !quiz.questions.length) throw new Error("Invalid quiz");
      quizIds.add(quiz.id);
      quiz.questions.forEach((question) => {
        if (!question.prompt) throw new Error("Missing question");
        if (question.element) {
          const element = data.elements?.[question.element];
          if (!element || element.symbol !== question.element || !element.name || !Number.isInteger(element.period) || !element.group || !["s", "p", "d", "f"].includes(element.block) || !element.subshell || !element.fact || !element.uses) throw new Error("Invalid element facts");
        }
        if (question.type === "numeric") {
          if (!Number.isFinite(question.answer) || (question.tolerance !== undefined && (!Number.isFinite(question.tolerance) || question.tolerance < 0))) throw new Error("Invalid numeric answer");
        } else if (question.type !== "choice" || !Array.isArray(question.options) || question.options.length < 2 || !question.options.every((option) => typeof option === "string" && option.length > 0) || !Number.isInteger(question.answer) || question.answer < 0 || question.answer >= question.options.length) {
          throw new Error("Invalid answer choices");
        }
      });
    });
  });
  return data;
}

if (typeof module !== "undefined" && module.exports) module.exports = { gradeQuiz, validatePracticeData };

if (typeof document !== "undefined") {
  const content = document.getElementById("practiceContent");
  let courses = [];
  let elements = {};

  function el(tag, text, className) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (className) node.className = className;
    return node;
  }

  function courseUrl(course) {
    return `#course=${encodeURIComponent(course.id)}`;
  }

  function link(text, href, className) {
    const node = el("a", text, className);
    node.href = href;
    return node;
  }

  function focusHeading() {
    const heading = content.querySelector("h2");
    if (heading) {
      heading.tabIndex = -1;
      heading.focus();
    }
  }

  function renderHome() {
    content.append(el("h2", "Choose your course"));
    const grid = el("div", undefined, "course-grid");
    courses.forEach((course, index) => {
      const card = link("", courseUrl(course), "card course-card");
      card.append(el("span", `COURSE 0${index + 1}`, "course-number"), el("h3", course.title), el("p", course.subject), el("strong", "Explore topics →"));
      grid.append(card);
    });
    content.append(grid);
  }

  function renderCourse(course) {
    content.append(link("← All courses", "#", "back"), el("h2", course.title), el("p", "Choose a quiz topic to start practicing.", "muted"));
    const panel = el("section", undefined, "card");
    panel.append(el("h3", "Practice topics"));
    if (!course.quizzes.length) {
      panel.append(el("p", "No practice quizzes available yet."), el("p", "Your teacher will add quizzes as you cover new topics.", "muted"));
    } else {
      const list = el("ul", undefined, "topic-list");
      course.quizzes.forEach((quiz) => {
        const item = el("li");
        item.append(link(quiz.title, `${courseUrl(course)}&quiz=${encodeURIComponent(quiz.id)}`));
        list.append(item);
      });
      panel.append(list);
    }
    content.append(panel);
  }

  function answerLabel(question, answer) {
    if (answer === undefined || answer === null || String(answer).trim() === "") return "Not answered";
    return question.type === "numeric" ? `${answer}${question.unit ? ` ${question.unit}` : ""}` : question.options[Number(answer)];
  }

  function renderElementFacts(element) {
    const card = el("section", undefined, "element-facts");
    card.append(el("h3", `${element.name} (${element.symbol})`));
    const details = el("dl", undefined, "element-details");
    [["Atomic number", element.atomicNumber], ["Period", element.period], ["Group", element.group], ["Block", element.block], ["Block subshell", element.subshell]].forEach(([label, value]) => {
      details.append(el("dt", label), el("dd", String(value)));
    });
    card.append(details, el("p", "The block subshell identifies the s, p, d, or f region of the table; it is not always the outermost occupied subshell.", "muted"), el("h4", "Interesting fact"), el("p", element.fact), el("h4", "Uses and applications"), el("p", element.uses));
    if (element.source) {
      const source = link("Learn more: Royal Society of Chemistry ↗", element.source);
      source.target = "_blank";
      source.rel = "noopener noreferrer";
      card.append(source);
    }
    return card;
  }

  function renderResult(course, quiz, answers) {
    const result = gradeQuiz(quiz, answers);
    const summary = el("section", undefined, "card");
    summary.tabIndex = -1;
    summary.setAttribute("aria-label", "Final score");
    summary.append(el("h2", "Your results"), el("h3", quiz.title), el("p", `${result.score} out of ${result.total}`, "score"), el("p", `${result.percent}% · ${result.score} correct · ${result.total - result.score} incorrect`));
    const actions = el("div", undefined, "actions");
    const retry = el("button", "Try again");
    retry.type = "button";
    retry.addEventListener("click", () => { content.replaceChildren(); renderQuiz(course, quiz); focusHeading(); });
    actions.append(retry, link("Back to topics", courseUrl(course), "button secondary"));
    summary.append(actions);
    content.append(summary);
    summary.focus();
  }

  function renderQuiz(course, quiz) {
    content.append(link(`← ${course.title} topics`, courseUrl(course), "back"), el("h2", quiz.title), el("p", `${course.title} · ${quiz.questions.length} questions · Choose an answer to see feedback immediately. Your first answer counts. Your final score appears automatically when you finish.`, "muted"));
    const form = el("form");
    const answers = Array(quiz.questions.length).fill(null);
    const progress = el("p", `0 of ${quiz.questions.length} answered`, "muted");
    progress.setAttribute("role", "status");
    content.append(progress);
    form.addEventListener("submit", (event) => event.preventDefault());
    quiz.questions.forEach((question, index) => {
      const fieldset = el("fieldset");
      fieldset.append(el("legend", `Question ${index + 1}`), el("p", question.prompt, "prompt"));
      const feedback = el("div");
      feedback.setAttribute("role", "status");
      function answerQuestion(value) {
        if (answers[index] !== null || String(value).trim() === "") return;
        answers[index] = value;
        const result = gradeQuiz(quiz, answers).results[index];
        fieldset.querySelectorAll("input").forEach((input) => { input.disabled = true; });
        feedback.className = `feedback${result.correct ? " correct" : ""}`;
        feedback.append(el("strong", result.correct ? "Correct!" : "Incorrect."), el("p", `Correct answer: ${answerLabel(question, question.answer)}`));
        if (question.explanation) feedback.append(el("p", question.explanation));
        if (result.correct && elements[question.element]) feedback.append(renderElementFacts(elements[question.element]));
        const answered = answers.filter((answer) => answer !== null).length;
        progress.textContent = `${answered} of ${quiz.questions.length} answered`;
        if (answered === quiz.questions.length) renderResult(course, quiz, answers);
      }
      if (question.type === "numeric") {
        const label = el("label", `Your answer${question.unit ? ` (${question.unit})` : ""}`);
        const input = el("input");
        input.type = "number";
        input.step = "any";
        input.name = `q${index}`;
        input.required = true;
        input.addEventListener("change", () => {
          if (input.validity.valid) answerQuestion(input.value);
        });
        label.append(input);
        fieldset.append(label);
      } else {
        question.options.forEach((option, optionIndex) => {
          const label = el("label", undefined, "option");
          const input = el("input");
          input.type = "radio";
          input.name = `q${index}`;
          input.value = String(optionIndex);
          input.required = true;
          input.addEventListener("change", () => answerQuestion(input.value));
          label.append(input, el("span", option));
          fieldset.append(label);
        });
      }
      fieldset.append(feedback);
      form.append(fieldset);
    });
    content.append(form);
  }

  function render() {
    const params = new URLSearchParams(location.hash.slice(1));
    const course = courses.find((item) => item.id === params.get("course"));
    const quiz = course?.quizzes.find((item) => item.id === params.get("quiz"));
    content.replaceChildren();
    if (quiz) renderQuiz(course, quiz);
    else if (course) renderCourse(course);
    else renderHome();
    focusHeading();
  }

  async function load() {
    try {
      const response = await fetch(`practice-data.json?v=${Date.now()}`, { cache: "no-store" });
      if (!response.ok) throw new Error("Unable to load quizzes");
      const data = validatePracticeData(await response.json());
      courses = data.courses;
      elements = data.elements || {};
      window.addEventListener("hashchange", render);
      render();
    } catch (_error) {
      content.replaceChildren(el("h2", "Practice is temporarily unavailable"), el("p", "Please reload the page to try again."));
      const retry = el("button", "Reload practice");
      retry.type = "button";
      retry.addEventListener("click", load);
      content.append(retry);
    }
  }
  load();
}
