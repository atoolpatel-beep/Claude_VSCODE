const STORAGE_KEY = "todo-app.items";

const form = document.getElementById("add-form");
const input = document.getElementById("new-todo");
const dateInput = document.getElementById("new-date");
const list = document.getElementById("list");
const empty = document.getElementById("empty");
const count = document.getElementById("count");
const clearDone = document.getElementById("clear-done");
const exportButton = document.getElementById("export");
const template = document.getElementById("todo-template");
const filterButtons = document.querySelectorAll(".filter");

let todos = load();
let filter = "all";

function load() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) ?? [];
  } catch {
    return [];
  }
}

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(todos));
  } catch {
    // Storage unavailable (e.g. private mode); keep working in memory.
  }
}

const dateFormat = new Intl.DateTimeFormat(undefined, {
  weekday: "short",
  month: "short",
  day: "numeric",
  year: "numeric",
});

// Local date as "YYYY-MM-DD", matching the value format of <input type="date">.
function todayKey() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

function renderDate(el, due) {
  el.hidden = !due;
  if (!due) return;

  const [year, month, day] = due.split("-").map(Number);
  const today = todayKey();
  const label = due === today ? "Today" : due < today ? "Overdue" : "Due";

  el.dateTime = due;
  el.textContent = `${label} · ${dateFormat.format(new Date(year, month - 1, day))}`;
  el.classList.toggle("is-today", due === today);
  el.classList.toggle("is-overdue", due < today);
}

function render() {
  const visible = todos.filter((t) =>
    filter === "active" ? !t.done : filter === "done" ? t.done : true
  );

  list.replaceChildren(
    ...visible.map((todo) => {
      const item = template.content.firstElementChild.cloneNode(true);
      item.dataset.id = todo.id;
      item.classList.toggle("is-done", todo.done);
      item.querySelector("input").checked = todo.done;
      item.querySelector(".todo__text").textContent = todo.text;
      renderDate(item.querySelector(".todo__date"), todo.due);
      return item;
    })
  );

  const left = todos.filter((t) => !t.done).length;
  count.textContent = `${left} task${left === 1 ? "" : "s"} left`;
  empty.hidden = visible.length > 0;
  clearDone.hidden = !todos.some((t) => t.done);
  exportButton.disabled = todos.length === 0;
}

// Tasks saved before `created` was recorded: ids start with Date.now() in base 36.
function createdAt(todo) {
  if (todo.created) return todo.created;
  const ms = parseInt(todo.id.slice(0, 8), 36);
  return ms > Date.UTC(2020, 0, 1) && ms <= Date.now() ? ms : null;
}

function update(next) {
  todos = next;
  save();
  render();
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  const text = input.value.trim();
  if (!text) return;
  const id = Date.now().toString(36) + Math.random().toString(36).slice(2);
  update([...todos, { id, text, due: dateInput.value, done: false, created: Date.now() }]);
  input.value = "";
  dateInput.value = "";
  input.focus();
});

list.addEventListener("change", (event) => {
  const id = event.target.closest(".todo").dataset.id;
  const done = event.target.checked;
  update(
    todos.map((t) => (t.id === id ? { ...t, done, completedAt: done ? Date.now() : null } : t))
  );
});

list.addEventListener("click", (event) => {
  if (!event.target.matches(".todo__delete")) return;
  const id = event.target.closest(".todo").dataset.id;
  update(todos.filter((t) => t.id !== id));
});

clearDone.addEventListener("click", () => update(todos.filter((t) => !t.done)));

// Exports every task, regardless of the active filter.
exportButton.addEventListener("click", () =>
  exportTodosToExcel(
    todos.map((t) => ({ ...t, created: createdAt(t) })),
    todayKey()
  )
);

filterButtons.forEach((button) =>
  button.addEventListener("click", () => {
    filter = button.dataset.filter;
    filterButtons.forEach((b) => b.classList.toggle("is-active", b === button));
    render();
  })
);

document.getElementById("today").textContent = new Intl.DateTimeFormat(undefined, {
  weekday: "long",
  month: "long",
  day: "numeric",
}).format(new Date());

render();
