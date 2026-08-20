const NAVY = "#0a2540";
const LINE_COLORS = ["rgb(150, 175, 209)", "#667788", NAVY];
const charts = [];

const monthFormatter = new Intl.DateTimeFormat("es-AR", { month: "short", year: "2-digit", timeZone: "UTC" });
const fullDateFormatter = new Intl.DateTimeFormat("es-AR", { day: "numeric", month: "short", year: "numeric", timeZone: "UTC" });

function asDate(value) {
  return new Date(`${value}T00:00:00Z`);
}

function inferPercentScale(series) {
  if (series.format !== "percent") return 1;
  const values = series.data.slice(-12).map(point => Math.abs(point.value));
  return values.length && Math.max(...values) <= 1 ? 100 : 1;
}

function valueFormatter(series, compact = false) {
  const scale = inferPercentScale(series);
  if (series.format === "percent") return value => `${(value * scale).toLocaleString("es-AR", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
  if (series.format === "currency") return value => `$ ${value.toLocaleString("es-AR", { maximumFractionDigits: 0, notation: compact ? "compact" : "standard" })}`;
  if (series.format === "ars_millions") return value => `$ ${value.toLocaleString("es-AR", { maximumFractionDigits: 0, notation: compact ? "compact" : "standard" })} M`;
  if (series.format === "usd_millions") return value => `USD ${value.toLocaleString("es-AR", { maximumFractionDigits: 0, notation: compact ? "compact" : "standard" })}`;
  if (series.format === "exchange_rate") return value => `$ ${value.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  return value => value.toLocaleString("es-AR", { maximumFractionDigits: 1, notation: compact ? "compact" : "standard" });
}

function initialWindow(series) {
  const limits = { day: 365, month: 120, quarter: 48, semester: 30, year: 20 };
  return Math.min(series.data.length, limits[series.frequency] || 120);
}

function renderSummary(seriesList) {
  const selectedCodes = ["P1", "A2", "S1", "E1", "X1"];
  const selected = selectedCodes.map(code => seriesList.find(item => item.code === code)).filter(Boolean);
  document.getElementById("summary").innerHTML = selected.map(series => {
    const latest = series.data.at(-1);
    return `<article class="summary-card">
      <h2>${series.title}</h2>
      <p class="summary-value">${valueFormatter(series, true)(latest.value)}</p>
      <p class="summary-date">${fullDateFormatter.format(asDate(latest.date))}</p>
    </article>`;
  }).join("");
}

function createChartCard(series) {
  const article = document.createElement("article");
  article.className = "chart-card";
  article.dataset.group = series.group;
  const latest = series.lines ? series.lines.at(-1).data.at(-1) : series.data.at(-1);
  article.innerHTML = `<h2>${series.title}</h2>
    <p class="subtitle">${series.subtitle}</p>
    <div class="chart-wrap"><canvas aria-label="${series.title}" role="img"></canvas></div>
    <div class="time-control">
      <label for="window-${series.code}">Ventana temporal</label>
      <input id="window-${series.code}" type="range" min="2" max="${series.data.length}" value="${initialWindow(series)}" step="1">
      <span class="window-label" aria-live="polite"></span>
    </div>
    <div class="chart-meta"><span>Ultimo: ${valueFormatter(series)(latest.value)}</span><span>${fullDateFormatter.format(asDate(latest.date))}</span></div>`;
  let data = series.data.slice(-initialWindow(series));
  const startDate = () => data[0].date;
  const datasets = series.lines
    ? series.lines.map((line, index) => ({
        label: line.label,
        data: line.data.filter(point => point.date >= startDate()).map(point => ({ x: point.date, y: point.value })),
        borderColor: line.color || LINE_COLORS[index], borderWidth: 2.75, pointRadius: 0, pointHoverRadius: 3, tension: .18
      }))
    : [{ data: data.map(point => point.value), borderColor: NAVY, borderWidth: 2.75, pointRadius: 0, pointHoverRadius: 3, tension: .18 }];
  const format = valueFormatter(series, true);
  const chart = new Chart(article.querySelector("canvas"), {
    type: "line",
    data: {
      labels: data.map(point => point.date),
      datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: Boolean(series.lines), labels: { color: NAVY, boxWidth: 24, boxHeight: 2 } },
        tooltip: { displayColors: Boolean(series.lines), callbacks: { title: items => fullDateFormatter.format(asDate(items[0].label)), label: item => `${item.dataset.label ? `${item.dataset.label}: ` : ""}${format(item.parsed.y)}` } }
      },
      scales: {
        x: {
          grid: { display: false },
          border: { display: true, color: NAVY, width: 1 },
          ticks: { color: NAVY, maxTicksLimit: 7, maxRotation: 0, callback: (_, index) => monthFormatter.format(asDate(data[index].date)) }
        },
        y: {
          grid: { display: false },
          border: { display: false },
          ticks: { color: NAVY, maxTicksLimit: 5, callback: value => format(value) }
        }
      }
    }
  });
  const slider = article.querySelector("input[type=range]");
  const windowLabel = article.querySelector(".window-label");
  const updateWindow = () => {
    const count = Number(slider.value);
    data = series.data.slice(-count);
    chart.data.labels = data.map(point => point.date);
    if (series.lines) {
      chart.data.datasets.forEach((dataset, index) => {
        dataset.data = series.lines[index].data.filter(point => point.date >= data[0].date).map(point => ({ x: point.date, y: point.value }));
      });
    } else {
      chart.data.datasets[0].data = data.map(point => point.value);
    }
    windowLabel.textContent = `${monthFormatter.format(asDate(data[0].date))} - ${monthFormatter.format(asDate(data.at(-1).date))}`;
    chart.update("none");
  };
  slider.addEventListener("input", updateWindow);
  updateWindow();
  charts.push(chart);
  return article;
}

function renderFilters(seriesList) {
  const groups = ["Todos", ...new Set(seriesList.map(item => item.group))];
  const nav = document.getElementById("filters");
  nav.innerHTML = groups.map((group, index) => `<button class="filter-button${index === 0 ? " active" : ""}" data-filter="${group}">${group}</button>`).join("");
  nav.addEventListener("click", event => {
    const button = event.target.closest("button");
    if (!button) return;
    nav.querySelectorAll("button").forEach(item => item.classList.toggle("active", item === button));
    document.querySelectorAll(".chart-card").forEach(card => {
      card.hidden = button.dataset.filter !== "Todos" && card.dataset.group !== button.dataset.filter;
    });
  });
}

async function init() {
  try {
    const response = await fetch(`data/series.json?v=${Date.now()}`);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    renderSummary(payload.series);
    renderFilters(payload.series);
    const container = document.getElementById("charts");
    payload.series.forEach(series => container.appendChild(createChartCard(series)));
    const updated = new Date(payload.updated_at);
    document.getElementById("updated").textContent = `Actualizado ${updated.toLocaleString("es-AR", { dateStyle: "medium", timeStyle: "short" })}`;
    if (payload.errors?.length) {
      const warning = document.getElementById("warning");
      warning.hidden = false;
      warning.textContent = `La ultima actualizacion mantuvo datos anteriores en ${payload.errors.length} serie(s).`;
    }
  } catch (error) {
    document.getElementById("updated").textContent = "No se pudieron cargar los datos";
    const warning = document.getElementById("warning");
    warning.hidden = false;
    warning.textContent = `Error de carga: ${error.message}`;
  }
}

init();
