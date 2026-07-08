const palette = ["#175d9c", "#25a36f", "#f2b705", "#d94f45", "#6b7a90", "#7a5cff", "#00a6a6"];
const data = window.dashboardData || {};

function chart(id, type, labels, values, options = {}) {
  const el = document.getElementById(id);
  if (!el) return;
  new Chart(el, {
    type,
    data: {
      labels,
      datasets: [{
        label: options.label || "%",
        data: values,
        backgroundColor: type === "doughnut" ? palette : values.map(v => v >= 85 ? "#25a36f" : v >= 70 ? "#f2b705" : "#d94f45"),
        borderColor: "#175d9c",
        tension: .3
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      normalized: true,
      indexAxis: options.horizontal ? "y" : "x",
      scales: type === "doughnut" ? {} : {
        x: { ticks: { autoSkip: true, maxTicksLimit: 12 } },
        y: { beginAtZero: true, max: options.max || undefined, ticks: { autoSkip: true, maxTicksLimit: 10 } }
      },
      plugins: { legend: { display: type === "doughnut" } }
    }
  });
}

chart("factorChart", "bar", data.factores?.labels || [], data.factores?.values || [], { horizontal: true, max: 100 });
chart("preguntaChart", "bar", data.preguntas?.labels || [], data.preguntas?.values || [], { max: 100 });
chart("sexoChart", "doughnut", data.sexo?.labels || [], data.sexo?.values || []);
chart("escolaridadChart", "doughnut", data.escolaridad?.labels || [], data.escolaridad?.values || []);
chart("hotelChart", "bar", data.hoteles?.labels || [], data.hoteles?.values || [], { max: 100 });
chart("deptoChart", "bar", data.departamentos?.labels || [], data.departamentos?.values || [], { max: 100 });
chart("edadChart", "bar", data.edad?.labels || [], data.edad?.values || [], { label: "Encuestas" });
chart("antiguedadChart", "bar", data.antiguedad?.labels || [], data.antiguedad?.values || [], { label: "Encuestas" });
chart("campanaChart", "line", data.campanas?.labels || [], data.campanas?.values || [], { max: 100 });
