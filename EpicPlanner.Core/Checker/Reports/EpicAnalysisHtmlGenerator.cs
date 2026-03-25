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
        string firstSprint   = m.SprintLabels.FirstOrDefault() ?? "?";
        string lastSprint    = m.SprintLabels.LastOrDefault()  ?? "?";
        string currentSprint = string.IsNullOrWhiteSpace(m.EpicConsumptionSprintLabel)
            ? lastSprint : m.EpicConsumptionSprintLabel;
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

<div style="display:flex;justify-content:flex-end;margin-bottom:6px;gap:8px">
  <button class="chart-btn" onclick="expandAll()" style="padding:5px 12px;font-size:12px">&#9660; Tout d&eacute;plier</button>
  <button class="chart-btn" onclick="collapseAll()" style="padding:5px 12px;font-size:12px">&#9650; Tout r&eacute;duire</button>
</div>

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

<div class="section-title section-toggle" onclick="toggleSection('sec-consumption')">&#9889; Consommation {{currentSprint}} &mdash; donn&eacute;es Redmine</div>
<div id="sec-consumption" class="collapsible">
  <table class="epic-table" id="tbl-consumption"></table>
</div>

<div class="section-title section-toggle" onclick="toggleSection('sec-devstats')">&#128101; R&eacute;alisateurs &mdash; Planifi&eacute; vs Consomm&eacute; par sprint</div>
<div id="sec-devstats" class="collapsible">
  <table class="epic-table" id="tbl-devstats"></table>
</div>

<div class="section-title section-toggle" onclick="toggleSection('sec-heatmap')">&#128202; Heatmap des allocations planifi&eacute;es par sprint (toutes les epics)</div>
<div id="sec-heatmap" class="collapsible">
<table class="hm-table" id="heatmap"></table>
<div class="legend-row">
  <div class="leg"><div class="leg-box" style="background:#c8e6c9;border:1px solid #aed6b1"></div>&gt; 80h planifi&eacute;es</div>
  <div class="leg"><div class="leg-box" style="background:#fff9c4;border:1px solid #f0e68c"></div>30&ndash;80h</div>
  <div class="leg"><div class="leg-box" style="background:#ffe0b2;border:1px solid #ffcc80"></div>&lt; 30h</div>
  <div class="leg"><div class="leg-box" style="background:#ffcdd2;outline:2px solid #e53935;outline-offset:-2px"></div>Absent (alerte)</div>
  <div class="leg"><div class="leg-box" style="background:#e8f5e9;border:1px solid #a5d6a7"></div>Epic termin&eacute;e ce sprint</div>
  <div class="leg"><div class="leg-box" style="background:#f9f9f9;border:1px solid #ddd"></div>Non active</div>
</div>
</div>

<div class="footnote">Sources&nbsp;: <em>Planification_des_Epics.xlsx</em> &middot; <em>Sprint_XX_vN.xlsx</em> ({{firstSprint}}&ndash;{{lastSprint}}, derni&egrave;res versions) &middot; Graphiques n&eacute;cessitent une connexion internet (Chart.js CDN)</div>
</div>

<div class="modal-overlay" id="dev-modal" onclick="closeDevModal(event)">
  <div class="modal" id="dev-modal-box">
    <button class="modal-close" onclick="closeDevModal()">&#10005;</button>
    <h2 id="dev-modal-title"></h2>
    <div class="meta-row" id="dev-modal-meta"></div>
    <div class="chart-container"><canvas id="dev-modal-chart"></canvas></div>
    <div class="modal-stats" id="dev-modal-stats"></div>
    <div id="dev-chart-legend" style="margin-top:10px;font-size:11px;color:#555;display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:6px 10px;background:#f8f9ff;border-radius:6px;border:1px solid #e8eaf6"></div>
  </div>
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
        sb.Append(BuildConsumptionJs(m));

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
//  CONSUMPTION (current sprint, per epic, Redmine)
// ═══════════════════════════════════════════════════
function renderConsumptions() {
  if (!EPIC_CONSUMPTIONS || EPIC_CONSUMPTIONS.length === 0) {
    document.getElementById('tbl-consumption').innerHTML =
      '<tbody><tr><td colspan="7" style="padding:10px;color:#888;font-style:italic">Aucune donnée Redmine disponible pour ce sprint.</td></tr></tbody>';
    return;
  }

  const header = `<thead><tr>
    <th style="min-width:170px">Epic</th>
    <th>Planifi&eacute; (h)</th>
    <th>Consomm&eacute; (h)</th>
    <th>Restant Redmine (h)</th>
    <th style="min-width:120px">Taux utilisation</th>
    <th>D&eacute;passement (h)</th>
    <th>Observation</th>
  </tr></thead>`;

  // Sort: under-consumed with allocation first, then overhead, then ok, then no plan
  const sorted = [...EPIC_CONSUMPTIONS].sort((a, b) => {
    const scoreA = consumptionScore(a), scoreB = consumptionScore(b);
    if (scoreA !== scoreB) return scoreA - scoreB;
    return a.epicId.localeCompare(b.epicId);
  });

  const rows = sorted.map(e => {
    const usagePct = e.usageRatePct;
    const isUnder  = e.planned > 0 && usagePct < 80;
    const isOver   = usagePct > 120;
    const rowCls   = isUnder ? 'row-red' : isOver ? 'row-orange' : e.planned > 0 ? 'row-green' : '';

    const bar = e.planned > 0
      ? `<div class="prog-wrap">
           <div class="prog-bar"><div class="prog-fill ${isUnder?'red':isOver?'orange':'green'}" style="width:${Math.min(usagePct,100)}%"></div></div>
           <span class="prog-pct" style="color:${isUnder?'#dc3545':isOver?'#fd7e14':'#28a745'}">${usagePct.toFixed(0)}%</span>
         </div>`
      : '<span style="color:#aaa;font-size:11px">—</span>';

    const ovStyle = e.overhead > 0 ? 'color:#dc3545;font-weight:700'
                  : e.overhead < 0 ? 'color:#28a745'
                  : 'color:#aaa';
    const ovText  = e.overhead > 0 ? '+'+e.overhead+'h' : e.overhead < 0 ? e.overhead+'h' : '—';

    const obs = consumptionObs(e);

    return `<tr class="${rowCls}">
      <td><strong>${e.epicId}</strong><br><span style="color:#555;font-size:11px">${e.epicName}</span></td>
      <td style="font-size:12px">${e.planned > 0 ? e.planned+'h' : '—'}</td>
      <td style="font-size:12px;font-weight:700">${e.consumed}h</td>
      <td style="font-size:12px">${e.remaining}h</td>
      <td>${bar}</td>
      <td style="font-size:12px;${ovStyle}">${ovText}</td>
      <td style="font-size:11px;color:#666">${obs}</td>
    </tr>`;
  }).join('');

  document.getElementById('tbl-consumption').innerHTML = header + '<tbody>' + rows + '</tbody>';
}

function consumptionScore(e) {
  if (e.planned <= 0) return 3;
  if (e.usageRatePct < 80) return 0;
  if (e.usageRatePct > 120) return 1;
  return 2;
}

function consumptionObs(e) {
  if (e.planned <= 0 && e.consumed > 0) return 'Heures saisies sans planification associée.';
  if (e.planned <= 0) return 'Non planifié dans ce sprint.';
  if (e.consumed === 0) return 'Aucune heure saisie — planification non consommée.';
  if (e.usageRatePct < 50) return 'Forte sous-consommation (&lt; 50% du planifié).';
  if (e.usageRatePct < 80) return 'Sous-consommation (&lt; 80% du planifié).';
  if (e.usageRatePct > 120) return 'Dépassement du planifié.';
  if (e.overhead > 0) return 'Restant Redmine supérieur au prévu — probable sous-estimation.';
  return '';
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

  const allocData    = [...e.allocation, 0];
  const consumedData = [...(e.consumed || []), null]; // null for 'Actuel' point
  const remData      = e.remaining.map(v => v === null ? null : v);

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
          backgroundColor: allocData.map(h => h > 0 ? 'rgba(66,133,244,0.45)' : 'rgba(200,200,200,0.15)'),
          borderColor:     allocData.map(h => h > 0 ? 'rgba(66,133,244,0.80)' : 'rgba(180,180,180,0.30)'),
          borderWidth: 1.5,
          yAxisID: 'y',
          order: 3,
        },
        {
          label: 'Consommé réel (h)',
          data: consumedData,
          backgroundColor: 'rgba(192,132,252,0.55)',
          borderColor: 'rgba(147,51,234,0.85)',
          borderWidth: 1.5,
          yAxisID: 'y',
          order: 4,
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
       <span style="display:inline-block;width:14px;height:14px;background:rgba(66,133,244,0.45);border:1px solid rgba(66,133,244,0.8);border-radius:2px"></span>
       Allocation planifiée
     </span>
     <span style="display:inline-flex;align-items:center;gap:5px">
       <span style="display:inline-block;width:14px;height:14px;background:rgba(192,132,252,0.55);border:1px solid rgba(147,51,234,0.85);border-radius:2px"></span>
       Consommé réel
     </span>
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
     <span style="color:#666;font-style:italic">— L'écart entre les deux lignes = heures planifiées non consommées.</span>`;
}

// ═══════════════════════════════════════════════════
//  DEV STATS — Planifié vs Consommé par réalisateur
// ═══════════════════════════════════════════════════
function computeDevStats() {
  const n = SPRINT_LABELS.length - 1; // number of historical sprints (exclude 'Actuel')
  const devMap = new Map();

  EPICS.forEach(e => {
    // Assigned field may be comma, slash, or semicolon separated
    const raw = (e.assigned || '').trim();
    const assignees = raw ? raw.split(/[,\/;]+/).map(s => s.trim()).filter(s => s) : [];
    if (assignees.length === 0) return;
    const share = 1 / assignees.length;

    assignees.forEach(dev => {
      if (!devMap.has(dev)) {
        devMap.set(dev, {
          name:     dev,
          planned:  new Array(n).fill(0),
          consumed: new Array(n).fill(null)
        });
      }
      const st = devMap.get(dev);
      for (let i = 0; i < n; i++) {
        const alloc = (e.allocation && e.allocation[i] != null) ? e.allocation[i] : 0;
        st.planned[i] = Math.round((st.planned[i] + alloc * share) * 10) / 10;
      }
      for (let i = 0; i < n; i++) {
        const cons = (e.consumed && e.consumed[i] != null) ? e.consumed[i] : null;
        if (cons !== null) {
          st.consumed[i] = Math.round(((st.consumed[i] ?? 0) + cons * share) * 10) / 10;
        }
      }
    });
  });

  return Array.from(devMap.values()).sort((a, b) => a.name.localeCompare(b.name));
}

let DEV_STATS = null;

function renderDevTable() {
  DEV_STATS = computeDevStats();
  const header = `<thead><tr>
    <th style="min-width:170px">R&eacute;alisateur</th>
    <th>Total planifi&eacute; (h)</th>
    <th>Total consomm&eacute; (h)</th>
    <th>Taux de consommation</th>
    <th>Sprints actifs</th>
    <th>Graphique</th>
  </tr></thead>`;

  const rows = DEV_STATS.map(d => {
    const totalPlan = Math.round(d.planned.reduce((s, v) => s + (v || 0), 0) * 10) / 10;
    const totalCons = Math.round(d.consumed.reduce((s, v) => s + (v ?? 0), 0) * 10) / 10;
    const rate = totalPlan > 0.01 ? Math.round(totalCons / totalPlan * 100) : (totalCons > 0 ? 999 : 0);
    const activeCount = d.planned.filter(v => v > 0).length;
    const rateColor = rate < 70 ? '#dc3545' : rate > 130 ? '#fd7e14' : '#28a745';
    const bar = totalPlan > 0
      ? `<div class="prog-wrap">
           <div class="prog-bar"><div class="prog-fill ${rate<70?'red':rate>130?'orange':'green'}" style="width:${Math.min(rate,100)}%"></div></div>
           <span class="prog-pct" style="color:${rateColor}">${rate}%</span>
         </div>`
      : '<span style="color:#aaa;font-size:11px">—</span>';
    return `<tr class="clickable" onclick="openDevModal('${d.name.replace(/'/g,"\\'")}')">
      <td><strong>${d.name}</strong></td>
      <td style="font-size:12px">${totalPlan > 0 ? totalPlan+'h' : '—'}</td>
      <td style="font-size:12px;font-weight:700">${totalCons > 0 ? totalCons+'h' : '—'}</td>
      <td>${bar}</td>
      <td style="font-size:12px;color:#666">${activeCount} sprint${activeCount!==1?'s':''}</td>
      <td><button class="chart-btn" onclick="event.stopPropagation();openDevModal('${d.name.replace(/'/g,"\\'")}')">📊 Voir</button></td>
    </tr>`;
  }).join('');

  document.getElementById('tbl-devstats').innerHTML = header + '<tbody>' + rows + '</tbody>';
}

let devChartInstance = null;

function openDevModal(devName) {
  const d = DEV_STATS && DEV_STATS.find(x => x.name === devName);
  if (!d) return;

  document.getElementById('dev-modal-title').textContent = '👤 ' + d.name + ' — Planifié vs Consommé';

  const n = SPRINT_LABELS.length - 1;
  const totalPlan = Math.round(d.planned.reduce((s, v) => s + (v || 0), 0) * 10) / 10;
  const totalCons = Math.round(d.consumed.reduce((s, v) => s + (v ?? 0), 0) * 10) / 10;
  const activeSprints = d.planned.filter(v => v > 0).length;
  const rate = totalPlan > 0.01 ? Math.round(totalCons / totalPlan * 100) : (totalCons > 0 ? 999 : 0);
  document.getElementById('dev-modal-meta').innerHTML =
    `<span>📅 ${activeSprints} sprints actifs</span>
     <span>📋 Sprints : ${SPRINT_LABELS[0] ?? '?'} → ${SPRINT_LABELS[n-1] ?? '?'}</span>`;

  document.getElementById('dev-modal-stats').innerHTML = `
    <div class="modal-stat"><div class="sv" style="color:#283593">${totalPlan}h</div><div class="sl">Total planifié</div></div>
    <div class="modal-stat"><div class="sv" style="color:#28a745">${totalCons}h</div><div class="sl">Total consommé</div></div>
    <div class="modal-stat"><div class="sv" style="color:${rate<70?'#dc3545':rate>130?'#fd7e14':'#28a745'}">${rate}%</div><div class="sl">Taux consommation</div></div>`;

  document.getElementById('dev-modal').classList.add('open');
  renderDevChart(d);
}

function closeDevModal(evt) {
  if (evt && evt.target !== document.getElementById('dev-modal')) return;
  document.getElementById('dev-modal').classList.remove('open');
  if (devChartInstance) { devChartInstance.destroy(); devChartInstance = null; }
}

function renderDevChart(d) {
  if (devChartInstance) { devChartInstance.destroy(); devChartInstance = null; }
  const ctx = document.getElementById('dev-modal-chart').getContext('2d');
  const n = SPRINT_LABELS.length - 1;
  const labels = SPRINT_LABELS.slice(0, n);

  const allVals = [...d.planned, ...d.consumed.map(v => v ?? 0)].filter(v => v > 0);
  const yMax = allVals.length ? Math.ceil(Math.max(...allVals) * 1.15 / 10) * 10 : 100;

  devChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Planifié (h)',
          data: d.planned,
          backgroundColor: d.planned.map(h => h > 0 ? 'rgba(66,133,244,0.45)' : 'rgba(200,200,200,0.15)'),
          borderColor:     d.planned.map(h => h > 0 ? 'rgba(66,133,244,0.80)' : 'rgba(180,180,180,0.30)'),
          borderWidth: 1.5,
          order: 2,
        },
        {
          label: 'Consommé réel (h)',
          data: d.consumed,
          backgroundColor: 'rgba(192,132,252,0.55)',
          borderColor:     'rgba(147,51,234,0.85)',
          borderWidth: 1.5,
          order: 1,
        },
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
              if (si < 0) return [];
              const plan = d.planned[si] ?? 0;
              const cons = d.consumed[si];
              if (cons === null || cons === undefined) return ['Consommé : données non disponibles'];
              if (plan > 0.01) {
                const r = Math.round(cons / plan * 100);
                if (r < 70)       return [`⚠️ Sous-consommation : ${r}% du planifié`];
                if (r > 130)      return [`⚠️ Dépassement : ${r}% du planifié`];
                return [`✓ ${r}% du planifié consommé`];
              }
              if (cons > 0) return ['Heures saisies sans planification'];
              return [];
            }
          }
        }
      },
      scales: {
        y: {
          type: 'linear',
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

  document.getElementById('dev-chart-legend').innerHTML =
    `<strong>Lecture :</strong>
     <span style="display:inline-flex;align-items:center;gap:5px">
       <span style="display:inline-block;width:14px;height:14px;background:rgba(66,133,244,0.45);border:1px solid rgba(66,133,244,0.8);border-radius:2px"></span>
       Planifié
     </span>
     <span style="display:inline-flex;align-items:center;gap:5px">
       <span style="display:inline-block;width:14px;height:14px;background:rgba(192,132,252,0.55);border:1px solid rgba(147,51,234,0.85);border-radius:2px"></span>
       Consommé réel
     </span>
     <span style="color:#666;font-style:italic">— Les heures sont proratisées en cas d'affectation multiple sur un epic.</span>`;
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

function expandAll() {
  document.querySelectorAll('.collapsible').forEach(el => {
    el.classList.remove('collapsed');
    el.style.maxHeight = el.scrollHeight + 9999 + 'px';
    const title = el.previousElementSibling;
    if (title) title.classList.remove('collapsed');
  });
}

function collapseAll() {
  document.querySelectorAll('.collapsible').forEach(el => {
    el.classList.add('collapsed');
    el.style.maxHeight = '0px';
    const title = el.previousElementSibling;
    if (title) title.classList.add('collapsed');
  });
}

document.querySelectorAll('.collapsible').forEach(el => {
  el.style.maxHeight = el.scrollHeight + 9999 + 'px';
});

document.addEventListener('keydown', e => { if (e.key === 'Escape') { closeModal(); closeDevModal(); } });

// ═══════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════
renderCards();
renderTables();
renderPipeline();
renderConsumptions();
renderDevTable();
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

    private static string BuildConsumptionJs(EpicAnalysisReportModel m)
    {
        var sb = new StringBuilder();
        sb.Append("const EPIC_CONSUMPTIONS = [\n");
        foreach (EpicConsumptionEntry e in m.EpicConsumptions)
        {
            sb.AppendLine(
                $"  {{epicId:'{JsStr(e.EpicId)}', epicName:'{JsStr(e.EpicName)}'," +
                $" planned:{Fmt(e.Planned)}, consumed:{Fmt(e.Consumed)}, remaining:{Fmt(e.Remaining)}," +
                $" overhead:{Fmt(e.Overhead)}, usageRatePct:{Fmt(e.UsageRatePct)}}},");
        }
        sb.Append("];\n\n");
        return sb.ToString();
    }

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

            string cons = "[" + string.Join(", ", e.Consumed.Select(c =>
                c.HasValue ? c.Value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture) : "null")) + "]";

            sb.AppendLine(
                $"{{id:'{JsStr(e.Id)}', name:'{JsStr(e.Name)}', manager:'{JsStr(e.Manager)}'," +
                $" assigned:'{JsStr(e.Assigned)}', state:'{JsStr(e.State)}', risk:'{JsStr(e.Risk)}'," +
                $" orig:{origJs}, cur:{e.CurrentRemaining.ToString("F0", System.Globalization.CultureInfo.InvariantCulture)}," +
                $" riskSince:'{JsStr(e.RiskSince)}', stateLabel:'{JsStr(e.StateLabel)}'," +
                $" riskDesc:'{JsStr(e.RiskDesc)}'," +
                $" allocation:{alloc}," +
                $" remaining:{rem}," +
                $" consumed:{cons}}},");
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

    private static string Fmt(double v) =>
        v.ToString("F1", System.Globalization.CultureInfo.InvariantCulture);
}
