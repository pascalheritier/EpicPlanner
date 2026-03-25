using System.Text;

namespace EpicPlanner.Core.Checker.Reports;

/// <summary>
/// Generates a self-contained HTML report from an <see cref="EpicAnalysisReportModel"/>,
/// matching the style of Analyse_Epics_Retard.html.
/// </summary>
public class EpicAnalysisHtmlGenerator
{
    public void Generate(EpicAnalysisReportModel model, string outputPath)
    {
        string html = BuildHtml(model);
        File.WriteAllText(outputPath, html, Encoding.UTF8);
    }

    private static string BuildHtml(EpicAnalysisReportModel m)
    {
        string firstSprint = m.SprintLabels.FirstOrDefault() ?? "?";
        string lastSprint  = m.SprintLabels.LastOrDefault()  ?? "?";
        string generated   = m.GeneratedAt.ToString("dd MMMM yyyy",
            System.Globalization.CultureInfo.GetCultureInfo("fr-FR"));

        var sb = new StringBuilder();

        sb.Append($$"""
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Analyse Avancement des Epics — Athena</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f6f9; color: #333; font-size: 13px; }
.page { max-width: 1440px; margin: 0 auto; padding: 24px; }

.header { background: linear-gradient(135deg, #1a237e 0%, #283593 100%); color: white; padding: 24px 28px; border-radius: 8px; margin-bottom: 24px; }
.header h1 { font-size: 22px; font-weight: 600; margin-bottom: 4px; }
.header .subtitle { opacity: .85; font-size: 13px; }
.header .meta { margin-top: 10px; font-size: 11px; opacity: .7; }

.cards { display: flex; gap: 14px; margin-bottom: 22px; flex-wrap: wrap; }
.card { flex: 1; min-width: 140px; background: white; border-radius: 8px; padding: 14px 18px; box-shadow: 0 1px 4px rgba(0,0,0,.1); border-left: 4px solid #ccc; }
.card.red    { border-left-color: #dc3545; }
.card.orange { border-left-color: #fd7e14; }
.card.green  { border-left-color: #28a745; }
.card.blue   { border-left-color: #6c757d; }
.card.done   { border-left-color: #adb5bd; }
.card .count { font-size: 30px; font-weight: 700; line-height: 1; }
.card.red    .count { color: #dc3545; }
.card.orange .count { color: #fd7e14; }
.card.green  .count { color: #28a745; }
.card.blue   .count { color: #6c757d; }
.card.done   .count { color: #adb5bd; }
.card .label { font-size: 10px; color: #666; margin-top: 4px; text-transform: uppercase; letter-spacing: .5px; font-weight: 600; }
.card .desc  { font-size: 11px; color: #888; margin-top: 4px; }

.section-title { font-size: 14px; font-weight: 700; color: #1a237e; margin: 22px 0 9px; padding-bottom: 6px; border-bottom: 2px solid #e8eaf6; display: flex; align-items: center; gap: 8px; }

.badge { display: inline-block; padding: 2px 7px; border-radius: 20px; font-size: 10px; font-weight: 700; white-space: nowrap; }
.badge-red    { background: #fde8ea; color: #c62828; }
.badge-orange { background: #fff3e0; color: #e65100; }
.badge-green  { background: #e8f5e9; color: #1b5e20; }
.badge-blue   { background: #e3f2fd; color: #0d47a1; }
.badge-grey   { background: #f5f5f5; color: #555; }
.badge-done   { background: #e8f5e9; color: #2e7d32; }

.epic-table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); margin-bottom: 6px; }
.epic-table th { background: #283593; color: white; padding: 9px 11px; text-align: left; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .4px; white-space: nowrap; }
.epic-table td { padding: 8px 11px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }
.epic-table tr:last-child td { border-bottom: none; }
.epic-table tr.clickable { cursor: pointer; transition: background .12s; }
.epic-table tr.clickable:hover td { background: #f0f4ff; }
.epic-table .row-red    td:first-child { border-left: 3px solid #dc3545; }
.epic-table .row-orange td:first-child { border-left: 3px solid #fd7e14; }
.epic-table .row-green  td:first-child { border-left: 3px solid #28a745; }
.epic-table .row-done   td:first-child { border-left: 3px solid #adb5bd; }

.prog-wrap { display: flex; align-items: center; gap: 6px; }
.prog-bar  { flex: 1; min-width: 70px; height: 9px; background: #e0e0e0; border-radius: 5px; overflow: hidden; }
.prog-fill { height: 100%; border-radius: 5px; }
.prog-fill.red    { background: #dc3545; }
.prog-fill.orange { background: #fd7e14; }
.prog-fill.green  { background: #28a745; }
.prog-fill.grey   { background: #adb5bd; }
.prog-pct  { font-size: 11px; font-weight: 700; min-width: 34px; text-align: right; }

.chart-btn { border: none; background: #e8eaf6; color: #283593; border-radius: 5px; padding: 3px 8px; font-size: 11px; cursor: pointer; font-weight: 600; transition: background .15s; }
.chart-btn:hover { background: #c5cae9; }

.hm-table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
.hm-table th { background: #283593; color: white; padding: 7px 9px; font-size: 10px; text-align: center; font-weight: 700; white-space: nowrap; }
.hm-table th.left { text-align: left; min-width: 240px; padding-left: 12px; }
.hm-table td { padding: 5px 7px; border: 1px solid #eee; text-align: center; font-size: 11px; font-weight: 600; }
.hm-table td.nm { text-align: left; font-weight: 600; padding-left: 12px; cursor: pointer; }
.hm-table td.nm:hover { background: #f0f4ff !important; }
.cell-none  { background: #f9f9f9; color: #ccc; font-weight: 400; }
.cell-low   { background: #ffe0b2; color: #bf360c; }
.cell-med   { background: #fff9c4; color: #827717; }
.cell-high  { background: #c8e6c9; color: #1b5e20; }
.cell-alert { background: #ffcdd2; color: #b71c1c; outline: 2px solid #e53935; outline-offset: -2px; }
.cell-done  { background: #e8f5e9; color: #2e7d32; }

.modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.55); z-index: 1000; align-items: center; justify-content: center; }
.modal-overlay.open { display: flex; }
.modal { background: white; border-radius: 10px; padding: 24px; width: 90%; max-width: 820px; max-height: 92vh; overflow-y: auto; box-shadow: 0 8px 40px rgba(0,0,0,.25); position: relative; }
.modal-close { position: absolute; top: 14px; right: 16px; border: none; background: #f5f5f5; border-radius: 50%; width: 30px; height: 30px; font-size: 16px; cursor: pointer; color: #555; display: flex; align-items: center; justify-content: center; }
.modal-close:hover { background: #e0e0e0; }
.modal h2 { font-size: 16px; color: #1a237e; margin-bottom: 4px; padding-right: 36px; }
.modal .meta-row { font-size: 12px; color: #666; margin-bottom: 16px; display: flex; gap: 16px; flex-wrap: wrap; }
.modal .meta-row span { display: flex; align-items: center; gap: 4px; }
.chart-container { position: relative; height: 340px; }
.modal-stats { display: flex; gap: 12px; margin-top: 14px; flex-wrap: wrap; }
.modal-stat { flex: 1; min-width: 120px; background: #f8f9ff; border-radius: 6px; padding: 10px 14px; }
.modal-stat .sv { font-size: 20px; font-weight: 700; }
.modal-stat .sl { font-size: 10px; color: #666; text-transform: uppercase; letter-spacing: .4px; }
.modal-note { margin-top: 12px; background: #fff8e1; border-left: 3px solid #ffc107; padding: 8px 12px; border-radius: 0 4px 4px 0; font-size: 12px; color: #555; }

.legend-row { display: flex; gap: 14px; flex-wrap: wrap; margin-top: 10px; padding: 8px 12px; background: white; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
.leg { display: flex; align-items: center; gap: 5px; font-size: 11px; color: #555; }
.leg-box { width: 15px; height: 15px; border-radius: 3px; flex-shrink: 0; }

.footnote { margin-top: 22px; font-size: 11px; color: #aaa; text-align: center; }

.section-toggle { cursor: pointer; user-select: none; }
.section-toggle::after { content: ' ▼'; font-size: 10px; }
.section-toggle.collapsed::after { content: ' ▶'; }
.collapsible { overflow: hidden; transition: max-height .3s ease; }
.collapsible.collapsed { max-height: 0 !important; }

@media print {
  body { background: white; }
  .chart-btn, .modal-overlay { display: none !important; }
  .page { padding: 10px; }
  [class^="cell-"] { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .header, .epic-table th, .hm-table th { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
}
</style>
</head>
<body>
<div class="page">

<div class="header">
  <h1>Analyse de l'avancement des Epics &mdash; {{firstSprint}} &agrave; {{lastSprint}}</h1>
  <div class="subtitle">Projet Athena &nbsp;&middot;&nbsp; Identification des epics en retard par rapport &agrave; la charge souhait&eacute;e</div>
  <div class="meta">Rapport g&eacute;n&eacute;r&eacute; le {{generated}} &nbsp;&middot;&nbsp; Sprint courant&nbsp;: {{lastSprint}} &nbsp;&middot;&nbsp; <em>Cliquer sur une ligne pour afficher le graphique d'&eacute;volution</em></div>
</div>

<div class="cards" id="cards"></div>

<div class="section-title section-toggle" onclick="toggleSection('sec-critical')">&#128308; Epics critiques &mdash; Retard significatif ou allocation arr&ecirc;t&eacute;e</div>
<div id="sec-critical" class="collapsible">
  <table class="epic-table" id="tbl-critical"></table>
</div>

<div class="section-title section-toggle" onclick="toggleSection('sec-watch')">&#128992; Epics &agrave; surveiller &mdash; Allocation insuffisante ou rythme en baisse</div>
<div id="sec-watch" class="collapsible">
  <table class="epic-table" id="tbl-watch"></table>
</div>

<div class="section-title section-toggle" onclick="toggleSection('sec-ok')">&#128994; Epics dans les temps</div>
<div id="sec-ok" class="collapsible">
  <table class="epic-table" id="tbl-ok"></table>
</div>

<div class="section-title section-toggle" onclick="toggleSection('sec-done')">&#9989; Epics termin&eacute;s ou &agrave; l'arr&ecirc;t (depuis {{firstSprint}})</div>
<div id="sec-done" class="collapsible">
  <table class="epic-table" id="tbl-done"></table>
</div>

<div class="section-title section-toggle" onclick="toggleSection('sec-pipeline')">&#128203; Pipeline &mdash; En analyse ou en attente de d&eacute;veloppement</div>
<div id="sec-pipeline" class="collapsible">
  <table class="epic-table" id="tbl-pipeline"></table>
</div>

<div class="section-title">&#128202; Heatmap des allocations planifi&eacute;es par sprint (toutes les epics)</div>
<table class="hm-table" id="heatmap"></table>
<div class="legend-row">
  <div class="leg"><div class="leg-box" style="background:#c8e6c9;border:1px solid #aed6b1"></div>&gt; 80h planifi&eacute;es</div>
  <div class="leg"><div class="leg-box" style="background:#fff9c4;border:1px solid #f0e68c"></div>30&ndash;80h</div>
  <div class="leg"><div class="leg-box" style="background:#ffe0b2;border:1px solid #ffcc80"></div>&lt; 30h</div>
  <div class="leg"><div class="leg-box" style="background:#ffcdd2;outline:2px solid #e53935;outline-offset:-2px"></div>Absent (alerte)</div>
  <div class="leg"><div class="leg-box" style="background:#e8f5e9;border:1px solid #a5d6a7"></div>Epic termin&eacute;e ce sprint</div>
  <div class="leg"><div class="leg-box" style="background:#f9f9f9;border:1px solid #ddd"></div>Non active</div>
</div>

<div class="footnote">Sources&nbsp;: <em>Planification_des_Epics.xlsx</em> &middot; <em>Sprint_XX_vN.xlsx</em> ({{firstSprint}}&ndash;{{lastSprint}}, derni&egrave;res versions) &middot; Graphiques n&eacute;cessitent une connexion internet (Chart.js CDN)</div>
</div>

<div class="modal-overlay" id="modal" onclick="closeModal(event)">
  <div class="modal" id="modal-box">
    <button class="modal-close" onclick="closeModal()">&#10005;</button>
    <h2 id="modal-title"></h2>
    <div class="meta-row" id="modal-meta"></div>
    <div class="chart-container"><canvas id="modal-chart"></canvas></div>
    <div class="modal-stats" id="modal-stats"></div>
    <div id="chart-bar-legend" style="margin-top:10px;font-size:11px;color:#555;display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:6px 10px;background:#f8f9ff;border-radius:6px;border:1px solid #e8eaf6"></div>
    <div class="modal-note" id="modal-note" style="display:none;margin-top:8px"></div>
  </div>
</div>

<script>
""");

        // Data
        sb.Append(BuildSprintLabelsJs(m));
        sb.Append(BuildEpicsJs(m));
        sb.Append(BuildPipelineJs(m));

        // All JS logic (rendering + chart)
        sb.Append($$"""
// ═══════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════
function pct(e) {
  if (!e.orig || e.orig === 0) return e.cur === 0 ? 100 : 0;
  return Math.round((e.orig - e.cur) / e.orig * 100);
}
function progColor(e) {
  if (e.risk === 'critical') return 'red';
  if (e.risk === 'watch')    return 'orange';
  if (e.risk === 'done')     return 'grey';
  return 'green';
}
function badgeHTML(risk) {
  const map = {critical:'badge-red', watch:'badge-orange', ok:'badge-green', done:'badge-done'};
  const lbl = {critical:'Critique', watch:'À surveiller', ok:'OK', done:'Terminé'};
  return `<span class="badge ${map[risk]}">${lbl[risk]}</span>`;
}
function cellClass(h, isAlert) {
  if (h === 0 || h === null) return isAlert ? 'cell-alert' : 'cell-none';
  if (h > 80)  return 'cell-high';
  if (h >= 30) return 'cell-med';
  return 'cell-low';
}

// ═══════════════════════════════════════════════════
//  TABLE HEADER
// ═══════════════════════════════════════════════════
function tableHeader(showRisk = true) {
  return `<thead><tr>
    <th style="min-width:170px">Epic</th>
    <th>Responsable</th>
    <th>Ressources assign&eacute;es</th>
    <th style="min-width:130px">Avancement</th>
    <th>Consomm&eacute;</th>
    <th>Restant</th>
    ${showRisk ? '<th>Statut</th><th style="min-width:110px">Depuis</th>' : ''}
    <th>Graphique</th>
  </tr></thead>`;
}

function epicRow(e, showRisk = true) {
  const p = pct(e);
  const consumed = e.orig ? (e.orig - e.cur) : 0;
  const cls = {critical:'row-red', watch:'row-orange', ok:'row-green', done:'row-done'}[e.risk];
  return `<tr class="clickable ${cls}" onclick="openModal('${e.id}')">
    <td><strong>${e.id}</strong><br><span style="color:#555;font-size:11px">${e.name}</span></td>
    <td style="font-size:12px">${e.manager}</td>
    <td style="font-size:11px;color:#666">${e.assigned}</td>
    <td>
      <div class="prog-wrap">
        <div class="prog-bar"><div class="prog-fill ${progColor(e)}" style="width:${Math.min(p,100)}%"></div></div>
        <span class="prog-pct" style="color:${p<30?'#dc3545':p<60?'#fd7e14':'#28a745'}">${p}%</span>
      </div>
    </td>
    <td style="font-size:12px">${consumed}h / ${e.orig??'?'}h</td>
    <td style="font-size:12px;font-weight:700">${e.cur}h</td>
    ${showRisk ? `<td>${badgeHTML(e.risk)}</td><td style="font-size:11px">${e.riskSince}</td>` : ''}
    <td><button class="chart-btn" onclick="event.stopPropagation();openModal('${e.id}')">📈 Voir</button></td>
  </tr>`;
}

// ═══════════════════════════════════════════════════
//  RENDER TABLES
// ═══════════════════════════════════════════════════
function renderTables() {
  const groups = {critical:[], watch:[], ok:[], done:[]};
  EPICS.forEach(e => groups[e.risk].push(e));

  document.getElementById('tbl-critical').innerHTML =
    tableHeader() + '<tbody>' + groups.critical.map(e=>epicRow(e)).join('') + '</tbody>';
  document.getElementById('tbl-watch').innerHTML =
    tableHeader() + '<tbody>' + groups.watch.map(e=>epicRow(e)).join('') + '</tbody>';
  document.getElementById('tbl-ok').innerHTML =
    tableHeader() + '<tbody>' + groups.ok.map(e=>epicRow(e)).join('') + '</tbody>';
  document.getElementById('tbl-done').innerHTML =
    tableHeader(false) + '<tbody>' + groups.done.map(e => {
      const p = pct(e);
      const consumed = e.orig ? (e.orig - e.cur) : '—';
      return `<tr class="clickable row-done" onclick="openModal('${e.id}')">
        <td><strong>${e.id}</strong><br><span style="color:#555;font-size:11px">${e.name}</span></td>
        <td style="font-size:12px">${e.manager}</td>
        <td style="font-size:11px;color:#666">${e.assigned}</td>
        <td>
          <div class="prog-wrap">
            <div class="prog-bar"><div class="prog-fill grey" style="width:${Math.min(p,100)}%"></div></div>
            <span class="prog-pct" style="color:#adb5bd">${p}%</span>
          </div>
        </td>
        <td style="font-size:12px">${consumed}h / ${e.orig??'?'}h</td>
        <td style="font-size:12px;font-weight:700">${e.cur}h</td>
        <td><button class="chart-btn" onclick="event.stopPropagation();openModal('${e.id}')">📈 Voir</button></td>
      </tr>`;
    }).join('') + '</tbody>';
}

// ═══════════════════════════════════════════════════
//  PIPELINE TABLE
// ═══════════════════════════════════════════════════
function renderPipeline() {
  const header = `<thead><tr>
    <th style="min-width:170px">Epic</th>
    <th>Manager</th>
    <th>Analyste</th>
    <th>&Eacute;tat</th>
    <th>Estimation (h)</th>
    <th>D&eacute;pendances</th>
    <th>Notes</th>
  </tr></thead>`;
  const rows = PIPELINE.map(p => `<tr>
    <td><strong>${p.id}</strong><br><span style="color:#555;font-size:11px">${p.name}</span></td>
    <td style="font-size:12px">${p.manager}</td>
    <td style="font-size:12px">${p.analyst}</td>
    <td style="font-size:11px">${p.state}</td>
    <td style="font-size:12px">${p.rough != null ? p.rough+'h' : '—'}</td>
    <td style="font-size:11px;color:#666">${p.deps}</td>
    <td style="font-size:11px;color:#888">${p.notes}</td>
  </tr>`).join('');
  document.getElementById('tbl-pipeline').innerHTML = header + '<tbody>' + rows + '</tbody>';
}

// ═══════════════════════════════════════════════════
//  SUMMARY CARDS
// ═══════════════════════════════════════════════════
function renderCards() {
  const g = {critical:0, watch:0, ok:0, done:0};
  EPICS.forEach(e => g[e.risk]++);
  const first = SPRINT_LABELS[0] ?? '?';
  const last  = SPRINT_LABELS[SPRINT_LABELS.length - 2] ?? '?'; // last before 'Actuel'
  document.getElementById('cards').innerHTML = `
    <div class="card red">   <div class="count">${g.critical}</div><div class="label">Critiques</div><div class="desc">Bloqu&eacute;es ou sans allocation</div></div>
    <div class="card orange"><div class="count">${g.watch}</div><div class="label">&Agrave; surveiller</div><div class="desc">Allocation insuffisante</div></div>
    <div class="card green"> <div class="count">${g.ok}</div><div class="label">Dans les temps</div><div class="desc">Progression conforme</div></div>
    <div class="card done">  <div class="count">${g.done}</div><div class="label">Termin&eacute;s / arr&ecirc;t&eacute;s</div><div class="desc">Depuis ${first}</div></div>
    <div class="card blue">  <div class="count">${EPICS.length}</div><div class="label">Total epics suivis</div><div class="desc">${first} &rarr; ${last}</div></div>`;
}

// ═══════════════════════════════════════════════════
//  HEATMAP
// ═══════════════════════════════════════════════════
function renderHeatmap() {
  const order = ['critical','watch','ok','done'];
  const sorted = [...EPICS].sort((a,b) => order.indexOf(a.risk) - order.indexOf(b.risk));
  const sprintCount = SPRINT_LABELS.length - 1; // exclude 'Actuel'

  const sprintHeaders = SPRINT_LABELS.slice(0, sprintCount)
    .map((s,i) => `<th>${s}<br><small style="font-weight:400;opacity:.75">${SPRINT_DATES[i]??''}</small></th>`).join('');

  let html = `<thead><tr>
    <th class="left">Epic</th>${sprintHeaders}
    <th>Consomm&eacute;</th><th>Restant</th><th>Statut</th>
  </tr></thead><tbody>`;

  sorted.forEach(e => {
    const riskColor = {critical:'#fde8ea', watch:'#fff3e0', ok:'', done:'#f5f5f5'}[e.risk];
    const idColor   = {critical:'#c62828', watch:'#e65100', ok:'#2e7d32', done:'#777'}[e.risk];
    const consumed  = e.orig ? (e.orig - e.cur) : '?';
    const lastActiveIdx = e.allocation.reduce((acc,h,i) => h > 0 ? i : acc, -1);

    const cells = e.allocation.slice(0, sprintCount).map((h, i) => {
      const isDone = e.state === 'done';
      const shouldAlert = h === 0 && !isDone && e.risk !== 'ok'
        && (lastActiveIdx >= 0 ? i <= lastActiveIdx + 1 && i > 0 : false);
      const isDoneThisSprint = isDone && h > 0 && e.remaining[i+1] === 0;
      if (isDoneThisSprint) return `<td class="cell-done">${h > 0 ? Math.round(h) + 'h' : '✓'}</td>`;
      const cls = shouldAlert ? 'cell-alert' : cellClass(h, false);
      return `<td class="${cls}">${h > 0 ? Math.round(h) + 'h' : '—'}</td>`;
    }).join('');

    html += `<tr>
      <td class="nm" style="background:${riskColor}" onclick="openModal('${e.id}')">
        <strong style="color:${idColor}">${e.id}</strong>
        <div style="font-size:10px;color:#666;font-weight:400">${e.name}</div>
      </td>${cells}
      <td style="font-size:12px">${consumed}h</td>
      <td style="font-size:12px;font-weight:700">${e.cur}h</td>
      <td>${badgeHTML(e.risk)}</td>
    </tr>`;
  });
  html += '</tbody>';
  document.getElementById('heatmap').innerHTML = html;
}

// ═══════════════════════════════════════════════════
//  MODAL + CHART
// ═══════════════════════════════════════════════════
let chartInstance = null;

function openModal(epicId) {
  const e = EPICS.find(x => x.id === epicId);
  if (!e) return;

  document.getElementById('modal-title').textContent = `${e.id} — ${e.name}`;

  const statusBadge = {critical:'🔴 Critique', watch:'🟠 À surveiller', ok:'🟢 OK', done:'✅ Terminé'}[e.risk];
  document.getElementById('modal-meta').innerHTML =
    `<span>👤 ${e.manager}</span>
     <span>🔧 ${e.assigned||'—'}</span>
     <span>${statusBadge}</span>
     <span>📋 ${e.stateLabel}</span>`;

  const consumed = e.orig ? (e.orig - e.cur) : 0;
  const p = pct(e);
  const sprintCount = SPRINT_LABELS.length - 1;
  document.getElementById('modal-stats').innerHTML = `
    <div class="modal-stat"><div class="sv" style="color:#283593">${e.orig??'?'}h</div><div class="sl">Estimé initial</div></div>
    <div class="modal-stat"><div class="sv" style="color:#28a745">${consumed}h</div><div class="sl">Consommé (${p}%)</div></div>
    <div class="modal-stat"><div class="sv" style="color:${p<40?'#dc3545':p<70?'#fd7e14':'#28a745'}">${e.cur}h</div><div class="sl">Restant actuel</div></div>
    <div class="modal-stat"><div class="sv" style="color:#555">${e.allocation.reduce((s,h)=>s+h,0).toFixed(0)}h</div><div class="sl">Total planifié ${SPRINT_LABELS[0]}–${SPRINT_LABELS[sprintCount-1]}</div></div>`;

  const noteEl = document.getElementById('modal-note');
  if (e.riskDesc) { noteEl.textContent = e.riskDesc; noteEl.style.display = 'block'; }
  else noteEl.style.display = 'none';

  document.getElementById('modal').classList.add('open');
  renderChart(e);
}

function closeModal(evt) {
  if (evt && evt.target !== document.getElementById('modal')) return;
  document.getElementById('modal').classList.remove('open');
  if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
}

function renderChart(e) {
  if (chartInstance) { chartInstance.destroy(); chartInstance = null; }
  const ctx = document.getElementById('modal-chart').getContext('2d');

  const allocData = [...e.allocation, 0];
  const remData   = e.remaining.map(v => v === null ? null : v);

  const expectedRem = remData.map((_, i) => {
    if (i === 0) return remData[0];
    const base  = remData[i - 1];
    const alloc = allocData[i - 1];
    if (base === null) return null;
    return Math.max(0, base - alloc);
  });

  const reEstIdx = [];
  for (let i = 1; i < remData.length; i++) {
    if (remData[i] !== null && remData[i-1] !== null && remData[i] > remData[i-1])
      reEstIdx.push(i);
  }
  const ptColor  = remData.map((v, i) => v === null ? 'transparent' : reEstIdx.includes(i) ? '#e53935' : '#1a237e');
  const ptRadius = remData.map((v, i) => v === null ? 0 : i === remData.length - 1 ? 7 : reEstIdx.includes(i) ? 6 : 4);

  const allVals = [...allocData, ...remData, ...expectedRem].filter(v => v !== null && v > 0);
  const yMax    = allVals.length ? Math.ceil(Math.max(...allVals) * 1.12 / 10) * 10 : 100;

  chartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: SPRINT_LABELS,
      datasets: [
        {
          label: 'Allocation planifiée (h)',
          data: allocData,
          backgroundColor: allocData.map(h => h > 0 ? 'rgba(66,133,244,0.50)' : 'rgba(200,200,200,0.18)'),
          borderColor:     allocData.map(h => h > 0 ? 'rgba(66,133,244,0.80)' : 'rgba(180,180,180,0.30)'),
          borderWidth: 1.5,
          yAxisID: 'y',
          order: 3,
        },
        {
          label: 'Restantes attendues (si tout consommé)',
          data: expectedRem,
          type: 'line',
          borderColor: 'rgba(230,100,0,0.85)',
          backgroundColor: 'transparent',
          borderWidth: 2,
          borderDash: [6, 4],
          pointRadius: expectedRem.map(v => v === null ? 0 : 3),
          pointBackgroundColor: 'rgba(230,100,0,0.85)',
          pointBorderColor:     'rgba(230,100,0,0.85)',
          pointHoverRadius: 6,
          fill: false,
          tension: 0.15,
          yAxisID: 'y',
          order: 2,
          spanGaps: false,
        },
        {
          label: 'Heures restantes (réelles)',
          data: remData,
          type: 'line',
          borderColor: '#1a237e',
          backgroundColor: 'rgba(26,35,126,0.07)',
          borderWidth: 2.5,
          pointBackgroundColor: ptColor,
          pointBorderColor: ptColor,
          pointRadius: ptRadius,
          pointHoverRadius: 8,
          fill: true,
          tension: 0.15,
          yAxisID: 'y',
          order: 1,
          spanGaps: false,
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top', labels: { font: { size: 12 }, boxWidth: 14 } },
        tooltip: {
          callbacks: {
            afterBody(items) {
              const si = items[0]?.dataIndex ?? -1;
              if (si < 1) return [];
              const lines = [];
              const remCur = remData[si];
              const expCur = expectedRem[si];
              if (remCur !== null && expCur !== null) {
                const gap = remCur - expCur;
                if (gap > 2)       lines.push(`▲ Retard sur allocation : +${gap.toFixed(0)}h non consommées`);
                else if (gap < -2) lines.push(`▼ En avance sur allocation : ${(-gap).toFixed(0)}h de plus`);
                else               lines.push('✓ Conforme à l\'allocation');
              }
              const remPrev = remData[si - 1];
              if (remCur !== null && remPrev !== null) {
                const delta = remCur - remPrev;
                if (delta > 0) lines.push(`⚠️ Re-estimation : +${delta.toFixed(0)}h`);
              }
              return lines;
            }
          }
        }
      },
      scales: {
        y: {
          type: 'linear',
          position: 'left',
          min: 0,
          max: yMax,
          title: { display: true, text: 'Heures', font: { size: 11 }, color: '#444' },
          ticks: { font: { size: 11 } },
          grid: { color: 'rgba(0,0,0,0.06)' }
        },
        x: { ticks: { font: { size: 11 } } }
      }
    }
  });

  document.getElementById('chart-bar-legend').innerHTML =
    `<strong>Lecture :</strong>
     <span style="display:inline-flex;align-items:center;gap:5px">
       <svg width="26" height="10"><line x1="0" y1="5" x2="26" y2="5" stroke="#1a237e" stroke-width="2.5"/></svg>
       Restantes réelles
     </span>
     <span style="display:inline-flex;align-items:center;gap:5px">
       <svg width="26" height="10"><line x1="0" y1="5" x2="26" y2="5" stroke="rgba(230,100,0,0.85)" stroke-width="2" stroke-dasharray="5,3"/></svg>
       Restantes attendues (si toute l'allocation avait été consommée)
     </span>
     <span style="display:inline-flex;align-items:center;gap:5px">
       <span style="color:#e53935;font-size:14px;line-height:1">●</span>
       Point rouge = re-estimation à la hausse
     </span>
     <span style="color:#666;font-style:italic">— L'écart entre les deux lignes indique les heures planifiées non consommées.</span>`;
}

// ═══════════════════════════════════════════════════
//  TOGGLE SECTIONS
// ═══════════════════════════════════════════════════
function toggleSection(id) {
  const el = document.getElementById(id);
  const title = el.previousElementSibling;
  el.classList.toggle('collapsed');
  title.classList.toggle('collapsed');
  if (!el.style.maxHeight || el.style.maxHeight === '0px') {
    el.style.maxHeight = el.scrollHeight + 'px';
  }
}

document.querySelectorAll('.collapsible').forEach(el => {
  el.style.maxHeight = el.scrollHeight + 9999 + 'px';
});

document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

// ═══════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════
renderCards();
renderTables();
renderPipeline();
renderHeatmap();

setTimeout(() => {
  document.querySelectorAll('.collapsible').forEach(el => {
    el.style.maxHeight = el.scrollHeight + 'px';
  });
}, 100);
</script>
</body>
</html>
""");

        return sb.ToString();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // JS data builders
    // ─────────────────────────────────────────────────────────────────────────

    private static string BuildSprintLabelsJs(EpicAnalysisReportModel m)
    {
        // SPRINT_LABELS = [...sprint labels, 'Actuel']  (n+1 entries, matches remaining array length)
        var labels = m.SprintLabels.Select(l => $"'{JsStr(l)}'").ToList();
        labels.Add("'Actuel'");
        string dates = "[" + string.Join(",", m.SprintDates.Select(d => $"'{JsStr(d)}'")) + "]";
        return $"const SPRINT_LABELS = [{string.Join(",", labels)}];\n"
             + $"const SPRINT_DATES  = {dates};\n\n";
    }

    private static string BuildEpicsJs(EpicAnalysisReportModel m)
    {
        var sb = new StringBuilder();
        sb.Append("const EPICS = [\n");
        foreach (EpicAnalysisEntry e in m.Epics)
        {
            string alloc = "[" + string.Join(", ", e.Allocation.Select(a =>
                a.ToString("F1", System.Globalization.CultureInfo.InvariantCulture))) + "]";
            string rem = "[" + string.Join(", ", e.Remaining.Select(r =>
                r.HasValue ? r.Value.ToString("F0", System.Globalization.CultureInfo.InvariantCulture) : "null")) + "]";
            string origJs = e.OriginalEstimate.HasValue
                ? e.OriginalEstimate.Value.ToString("F0", System.Globalization.CultureInfo.InvariantCulture)
                : "null";

            sb.AppendLine(
                $"{{id:'{JsStr(e.Id)}', name:'{JsStr(e.Name)}', manager:'{JsStr(e.Manager)}'," +
                $" assigned:'{JsStr(e.Assigned)}', state:'{JsStr(e.State)}', risk:'{JsStr(e.Risk)}'," +
                $" orig:{origJs}, cur:{e.CurrentRemaining.ToString("F0", System.Globalization.CultureInfo.InvariantCulture)}," +
                $" riskSince:'{JsStr(e.RiskSince)}', stateLabel:'{JsStr(e.StateLabel)}'," +
                $" riskDesc:'{JsStr(e.RiskDesc)}'," +
                $" allocation:{alloc}," +
                $" remaining:{rem}}},");
        }
        sb.Append("];\n\n");
        return sb.ToString();
    }

    private static string BuildPipelineJs(EpicAnalysisReportModel m)
    {
        var sb = new StringBuilder();
        sb.Append("const PIPELINE = [\n");
        foreach (PipelineEpicEntry p in m.Pipeline)
        {
            string rough = p.RoughEstimate.HasValue
                ? p.RoughEstimate.Value.ToString("F0", System.Globalization.CultureInfo.InvariantCulture)
                : "null";
            sb.AppendLine(
                $"{{id:'{JsStr(p.Id)}', name:'{JsStr(p.Name)}', manager:'{JsStr(p.Manager)}'," +
                $" analyst:'{JsStr(p.Analyst)}', state:'{JsStr(p.State)}'," +
                $" rough:{rough}, deps:'{JsStr(p.Dependencies)}', notes:'{JsStr(p.Notes)}'}},");
        }
        sb.Append("];\n\n");
        return sb.ToString();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Utility
    // ─────────────────────────────────────────────────────────────────────────

    private static string JsStr(string? s) =>
        string.IsNullOrEmpty(s) ? string.Empty
        : s.Replace("\\", "\\\\")
           .Replace("'", "\\'")
           .Replace("\r", "")
           .Replace("\n", " ");
}
