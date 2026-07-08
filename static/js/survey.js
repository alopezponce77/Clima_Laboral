const answers = Array.from(document.querySelectorAll(".survey-answer"));
const openAnswers = Array.from(document.querySelectorAll(".survey-open"));
const progress = document.getElementById("surveyProgress");
const names = [...new Set(answers.map(input => input.name))];
const totalItems = names.length + openAnswers.length;

function updateProgress() {
  const answered = names.filter(name => document.querySelector(`input[name="${name}"]:checked`)).length;
  const openAnswered = openAnswers.filter(input => input.value.trim().length > 0).length;
  const percent = Math.round(((answered + openAnswered) / totalItems) * 100);
  progress.style.width = `${percent}%`;
  progress.textContent = `${percent}%`;
}

answers.forEach(input => input.addEventListener("change", updateProgress));
openAnswers.forEach(input => input.addEventListener("input", updateProgress));
updateProgress();
