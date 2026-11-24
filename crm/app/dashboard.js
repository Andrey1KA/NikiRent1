"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { BarController, BarElement, CategoryScale, Chart, Legend, LinearScale, Tooltip } from "chart.js";
import * as XLSX from "xlsx";
import {
  apartmentsOf,
  formatMoney,
  formatQty,
  formatReading,
  monthsOf,
  parseWorkbook,
  previousOf,
  rowsForMonth,
  sumField,
  warningsFor,
} from "../lib/rent.mjs";

Chart.register(BarController, BarElement, CategoryScale, LinearScale, Tooltip, Legend);
Chart.defaults.font.family = '"Segoe UI", "Segoe UI Emoji", sans-serif';

const PALETTE = ["#0e6b62", "#c45c26", "#1f4e79", "#8a6a2f", "#5c6b73"];
const AXIS = {
  ticks: { color: "#6d645b" },
  grid: { color: "#efe7dc" },
  border: { display: false },
};

export default function Dashboard() {
  const [rows, setRows] = useState([]);
  const [month, setMonth] = useState("май");
  const [apartment, setApartment] = useState("");
  const [ready, setReady] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [notice, setNotice] = useState(null);
  const utilRef = useRef(null);
  const compareRef = useRef(null);
  const utilChart = useRef(null);
  const compareChart = useRef(null);

  useEffect(() => {
    let cancelled = false;
    async function loadDefault() {
      try {
        const response = await fetch("/data/kommunalka.xlsx");
        if (!response.ok) throw new Error("Не удалось открыть встроенную таблицу");
        const nextRows = parseWorkbook(XLSX, await response.arrayBuffer());
        if (cancelled) return;
        setRows(nextRows);
        const months = monthsOf(nextRows);
        setMonth(months.includes("май") ? "май" : months[0] || "");
        setApartment(apartmentsOf(nextRows)[0] || "");
      } catch (error) {
        if (!cancelled) setLoadError(error.message || "Не удалось открыть таблицу");
      } finally {
        if (!cancelled) setReady(true);
      }
    }
    loadDefault();
    return () => {
      cancelled = true;
    };
  }, []);

  const months = useMemo(() => monthsOf(rows), [rows]);
  const apartments = useMemo(() => apartmentsOf(rows), [rows]);
  const monthRows = useMemo(() => rowsForMonth(rows, month), [rows, month]);
  const selected = monthRows.find((row) => row.apartment === apartment) || null;

  const annotated = useMemo(
    () => monthRows.map((row) => ({ ...row, warnings: warningsFor(row, previousOf(rows, row)) })),
    [monthRows, rows],
  );
  const visibleWarnings = annotated.filter((row) => !apartment || row.apartment === apartment);
  const reviewCount = annotated.filter((row) => row.warnings.length > 0).length;
  const warningCount = visibleWarnings.reduce((sum, row) => sum + row.warnings.length, 0);

  const utilSeries = useMemo(
    () => monthRows.map((row) => ({ name: row.apartment, value: row.utilities })),
    [monthRows],
  );
  const compareSeries = useMemo(() => {
    return apartments.map((name) => ({
      name,
      april: rows.find((row) => row.apartment === name && row.month === "апрель")?.payout,
      may: rows.find((row) => row.apartment === name && row.month === "май")?.payout,
    }));
  }, [apartments, rows]);

  useEffect(() => () => {
    utilChart.current?.destroy();
    compareChart.current?.destroy();
  }, []);

  useEffect(() => {
    drawChart(utilRef, utilChart, utilSeries.every((item) => typeof item.value !== "number") ? null : {
      labels: utilSeries.map((item) => item.name),
      datasets: [
        {
          data: utilSeries.map((item) => (typeof item.value === "number" ? item.value : null)),
          backgroundColor: utilSeries.map((item, index) => (item.name === apartment ? PALETTE[0] : PALETTE[index % PALETTE.length])),
          borderRadius: 6,
          maxBarThickness: 28,
        },
      ],
    }, false);
  }, [utilSeries, apartment]);

  useEffect(() => {
    const faded = month === "апрель" ? "#0e6b62" : "#d9cfc3";
    const strong = month === "май" ? "#c45c26" : "#d9cfc3";
    drawChart(compareRef, compareChart, compareSeries.length === 0 ? null : {
      labels: compareSeries.map((item) => item.name),
      datasets: [
        {
          label: "Апрель",
          data: compareSeries.map((item) => (typeof item.april === "number" ? item.april : null)),
          backgroundColor: month === "апрель" ? "#0e6b62" : faded,
          borderRadius: 6,
          maxBarThickness: 18,
        },
        {
          label: "Май",
          data: compareSeries.map((item) => (typeof item.may === "number" ? item.may : null)),
          backgroundColor: month === "май" ? "#c45c26" : strong,
          borderRadius: 6,
          maxBarThickness: 18,
        },
      ],
    }, true);
  }, [compareSeries, month]);

  async function onFileChange(event) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const nextRows = parseWorkbook(XLSX, await file.arrayBuffer());
      const nextMonths = monthsOf(nextRows);
      const nextApartments = apartmentsOf(nextRows);
      setRows(nextRows);
      setMonth(nextMonths.includes(month) ? month : nextMonths[0] || "");
      setApartment(nextApartments.includes(apartment) ? apartment : nextApartments[0] || "");
      setLoadError("");
      setNotice({ tone: "ok", text: `Файл прочитан: ${nextApartments.length} кв., ${nextRows.length} строк` });
    } catch (error) {
      setNotice({ tone: "error", text: error.message || "Не удалось прочитать файл" });
    }
  }

  const resources = selected
    ? [
        ["Газ", formatReading(selected.gasReading), selected.gasUse, selected.gasCharge],
        ["Свет", formatReading(selected.lightReading), selected.lightUse, selected.lightCharge],
        ["Горячая вода", formatReading(selected.hot1, selected.hot2), selected.hotUse, selected.hotCharge],
        ["Холодная вода", formatReading(selected.cold1, selected.cold2), selected.coldUse, selected.coldCharge],
        ["Горводснабж", "нет своего счётчика", selected.hotUse, selected.supplyCharge],
      ]
    : [];
  const selectedWarnings = annotated.find((row) => row.apartment === apartment)?.warnings || [];

  return (
    <div className="wrap">
      <header className="top">
        <div>
          <p className="eyebrow">NikiRent</p>
          <h1>Коммуналка и выплаты</h1>
          <p className="lede">Пять квартир, апрель и май. Суммы читаются из файла, в код они не вписаны.</p>
          {loadError ? <p className="message error">{loadError}</p> : null}
        </div>
        <div className="upload">
          <label className="button" htmlFor="file">Загрузить файл</label>
          <input
            id="file"
            className="file-input"
            type="file"
            accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
            onChange={onFileChange}
          />
          <p className="hint">xlsx или csv. Файл в проекте не изменяется.</p>
          {notice ? <p className={`message ${notice.tone}`} role="status">{notice.text}</p> : null}
        </div>
      </header>

      <section className="panel" aria-labelledby="summary-title">
        <div className="panel-head">
          <div>
            <h2 id="summary-title">Сводка</h2>
            <p className="hint">Цифры за выбранный месяц по всем квартирам. Смена месяца меняет суммы и графики.</p>
          </div>
          <div className="filters">
            <label className="month-label">
              Месяц
              <select id="month" value={month} onChange={(event) => setMonth(event.target.value)} disabled={!months.length}>
                {months.length === 0 ? <option value="">Нет данных</option> : null}
                {months.map((name) => <option key={name} value={name}>{name}</option>)}
              </select>
            </label>
            <label className="month-label">
              Квартира
              <select id="apartment" value={apartment} onChange={(event) => setApartment(event.target.value)} disabled={!apartments.length}>
                <option value="">Все квартиры</option>
                {apartments.map((name) => <option key={name} value={name}>{name}</option>)}
              </select>
            </label>
          </div>
        </div>

        <div className="kpis">
          <Kpi label="Квартир" value={ready ? String(apartments.length) : "—"} />
          <Kpi label="Аренда" value={ready ? formatMoney(sumField(monthRows, "rent")) : "—"} accent />
          <Kpi label="Коммуналка" value={ready ? formatMoney(sumField(monthRows, "utilities")) : "—"} />
          <Kpi label="К собственникам" value={ready ? formatMoney(sumField(monthRows, "payout")) : "—"} accent />
          <Kpi label="На проверке" value={ready ? String(reviewCount) : "—"} alert />
        </div>

        <div className="charts">
          <div className="chart-box">
            <p className="hint">Коммуналка по квартирам, {month || "месяц"}</p>
            <div className="chart-wrap">
              <canvas ref={utilRef} hidden={utilSeries.every((item) => typeof item.value !== "number")} aria-label="Коммуналка по квартирам" />
              {ready && utilSeries.every((item) => typeof item.value !== "number") ? <p className="empty">Коммуналка за этот месяц — нет данных, ноль на график не ставится.</p> : null}
            </div>
          </div>
          <div className="chart-box">
            <p className="hint">Итого к выплате: апрель и май. Выбранный месяц ярче.</p>
            <div className="chart-wrap">
              <canvas ref={compareRef} aria-label="Начисления за апрель и май" />
            </div>
          </div>
        </div>
      </section>

      <section className="panel" aria-labelledby="card-title">
        <div className="panel-head">
          <div>
            <h2 id="card-title">{apartment || "Квартиры"}</h2>
            <p className="hint">{month ? `${capitalize(month)}. Пустое значение показано как «нет данных».` : "Период появится после загрузки файла."}</p>
          </div>
          {selected ? <Status complete={selected.complete} /> : null}
        </div>
        <div className="split">
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Квартира</th>
                  <th className="num">Аренда</th>
                  <th className="num">Коммуналка</th>
                  <th className="num">К выплате</th>
                  <th className="num">Фин. учет</th>
                </tr>
              </thead>
              <tbody>
                {annotated.map((row) => (
                  <tr
                    key={row.apartment}
                    className={row.apartment === apartment ? "clickable selected" : "clickable"}
                    onClick={() => setApartment(row.apartment)}
                  >
                    <td>{row.apartment}</td>
                    <td className="num">{formatMoney(row.rent)}</td>
                    <td className="num">{formatMoney(row.utilities)}</td>
                    <td className="num">{formatMoney(row.payout)}</td>
                    <td className="num">{formatMoney(row.fin)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {selected ? (
            <div>
              <dl className="money-list">
                <dt>Аренда</dt><dd>{formatMoney(selected.rent)}</dd>
                <dt>Коммуналка</dt><dd>{formatMoney(selected.utilities)}</dd>
                <dt>Доп. расходы</dt><dd>{formatMoney(selected.extra)}</dd>
                <dt>Перенос с прошлого месяца</dt><dd>{formatMoney(selected.carry)}</dd>
                <dt>Итого к выплате</dt><dd>{formatMoney(selected.payout)}</dd>
                <dt>Перечислено</dt><dd>{formatMoney(selected.transferred)}</dd>
                <dt>Долг / переплата</dt><dd>{formatMoney(selected.debt)}</dd>
                <dt>Фин. учет</dt><dd>{formatMoney(selected.fin)}</dd>
              </dl>
              {selectedWarnings.length > 0 ? (
                <ul className="warnings">
                  {selectedWarnings.map((text) => <li key={text}>{text}</li>)}
                </ul>
              ) : <p className="empty">Предупреждений по этой квартире нет.</p>}
            </div>
          ) : <p className="empty">Выберите квартиру в таблице или в фильтре.</p>}
        </div>
        {selected ? (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Ресурс</th>
                  <th>Показание</th>
                  <th className="num">Расход</th>
                  <th className="num">Начислено</th>
                </tr>
              </thead>
              <tbody>
                {resources.map(([name, reading, use, charge]) => (
                  <tr key={name}>
                    <td>{name}</td>
                    <td>{reading}</td>
                    <td className="num">{formatQty(use)}</td>
                    <td className="num">{formatMoney(charge)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : null}
      </section>

      <section className="panel" aria-labelledby="warn-title">
        <div className="panel-head">
          <div>
            <h2 id="warn-title">Что проверить</h2>
            <p className="hint">Пороги 300 кВт·ч и 20 м³ — повод посмотреть показание, а не доказательство ошибки.</p>
          </div>
          <p className="count-pill">{warningCount} {plural(warningCount, "предупреждение", "предупреждения", "предупреждений")}</p>
        </div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Квартира</th>
                <th>Причина</th>
              </tr>
            </thead>
            <tbody>
              {visibleWarnings.flatMap((row) => row.warnings.map((text) => (
                <tr key={`${row.apartment}-${text}`}>
                  <td>{row.apartment}</td>
                  <td>{text}</td>
                </tr>
              )))}
            </tbody>
          </table>
          {ready && warningCount === 0 ? <p className="empty">За выбранный срез предупреждений нет.</p> : null}
        </div>
      </section>
    </div>
  );
}

function Kpi({ label, value, accent = false, alert = false }) {
  const tone = alert ? "kpi alert" : accent ? "kpi accent" : "kpi";
  return (
    <div className={tone}>
      <p className="total">{value}</p>
      <p className="hint">{label}</p>
    </div>
  );
}

function Status({ complete }) {
  if (complete === "полная") return <p className="pill ok">полная</p>;
  if (complete === "неполная") return <p className="pill warn">неполная</p>;
  return <p className="pill muted">{complete || "нет данных"}</p>;
}

function drawChart(ref, slot, config, legend) {
  const canvas = ref.current;
  if (!canvas || !config) {
    slot.current?.destroy();
    slot.current = null;
    return;
  }
  if (slot.current) {
    slot.current.data.labels = config.labels;
    slot.current.data.datasets = config.datasets;
    slot.current.update("none");
    return;
  }
  slot.current = new Chart(canvas, {
    type: "bar",
    data: config,
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: legend, labels: { color: "#6d645b", boxWidth: 12 } },
        tooltip: {
          callbacks: {
            label(item) {
              const value = item.parsed.y;
              return ` ${item.dataset.label ? `${item.dataset.label}: ` : ""}${formatMoney(value)}`;
            },
          },
        },
      },
      scales: {
        x: { ...AXIS, ticks: { ...AXIS.ticks, color: "#1e1a16", font: { size: 11 } }, grid: { display: false } },
        y: { ...AXIS, beginAtZero: true },
      },
    },
  });
}

function capitalize(value) {
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : value;
}

function plural(count, one, few, many) {
  const mod10 = count % 10;
  const mod100 = count % 100;
  if (mod10 === 1 && mod100 !== 11) return one;
  if (mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14)) return few;
  return many;
}
