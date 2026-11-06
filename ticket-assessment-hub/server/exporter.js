'use strict';
// Builds CSV and XLSX exports from the shared export tables. Read-only: never mutates the store.
const ExcelJS = require('exceljs');
const core = require('./core');

const KINDS = {
  all: { label: 'full-workbook', sheets: ['Quiz Registry', 'Raw Responses', 'Evaluation Results', 'Score Breakdown', 'Areas for Improvement', 'Cohort Performance'] },
  registry: { label: 'quiz-registry', sheets: ['Quiz Registry'] },
  raw: { label: 'raw-responses', sheets: ['Raw Responses'] },
  evaluations: { label: 'evaluation-results', sheets: ['Evaluation Results'] },
  breakdown: { label: 'score-breakdown', sheets: ['Score Breakdown'] },
  areas: { label: 'areas-for-improvement', sheets: ['Areas for Improvement'] },
  cohort: { label: 'cohort-performance', sheets: ['Cohort Performance'] },
  trainee: { label: 'trainee-report', sheets: ['Evaluation Results', 'Score Breakdown', 'Areas for Improvement', 'Raw Responses'] },
};

function slug(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40);
}

async function build(db, kind, format, filter) {
  const spec = KINDS[kind];
  if (!spec) {
    const e = new Error('Unknown export type.');
    e.status = 400;
    throw e;
  }
  if (kind === 'trainee' && !filter.traineeId) {
    const e = new Error('Choose a trainee for the individual report.');
    e.status = 400;
    throw e;
  }
  const tables = core.exportTables(db, filter);
  const stamp = new Date().toISOString().slice(0, 10);
  let name = `${spec.label}`;
  if (kind === 'trainee') name += '-' + slug((db.trainees.find((t) => t.id === filter.traineeId) || {}).name || filter.traineeId);
  if (filter.cohort) name += '-' + slug(filter.cohort);
  name += `-${stamp}`;

  if (format === 'csv') {
    if (spec.sheets.length > 1 && kind !== 'trainee') {
      const e = new Error('CSV holds one table. Pick a single table, or export the full workbook as Excel.');
      e.status = 400;
      throw e;
    }
    // Trainee CSV: the per-criterion breakdown is the most useful single table.
    const sheet = kind === 'trainee' ? 'Score Breakdown' : spec.sheets[0];
    const csv = core.toCSV(tables[sheet]) || 'No rows match this export.\r\n';
    return { buffer: Buffer.from('﻿' + csv, 'utf8'), filename: `${name}.csv`, contentType: 'text/csv; charset=utf-8' };
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Ticket Assessment Hub (staging)';
  wb.created = new Date();
  if (kind === 'trainee') {
    const tr = db.trainees.find((t) => t.id === filter.traineeId) || {};
    const rows = core.assignmentRows(db).filter((r) => r.traineeId === filter.traineeId);
    const ws = wb.addWorksheet('Summary');
    ws.addRow(['Trainee report']).font = { bold: true, size: 14 };
    ws.addRow(['Trainee', tr.name || '']);
    ws.addRow(['Email', tr.email || '']);
    ws.addRow(['Cohort', tr.cohort || '']);
    ws.addRow(['Generated', new Date().toISOString()]);
    ws.addRow([]);
    const h = ws.addRow(['Quiz', 'Status', 'Points earned', 'Points possible', 'Score (%)', 'Score source', 'Result']);
    h.font = { bold: true };
    rows.forEach((r) => ws.addRow([r.quizTitle, r.status, r.earned ?? '', r.possible ?? '', r.percent ?? '', r.scoreSource === 'trainer' ? 'Trainer final' : r.scoreSource === 'ai' ? 'AI (not yet approved)' : '', r.result ?? '']));
    ws.columns.forEach((c) => (c.width = 24));
  }
  spec.sheets.forEach((sheetName) => {
    const rows = tables[sheetName];
    const ws = wb.addWorksheet(sheetName);
    if (!rows.length) {
      ws.addRow(['No rows match this export.']);
      return;
    }
    const headers = Object.keys(rows[0]);
    rows.forEach((r) => Object.keys(r).forEach((k) => !headers.includes(k) && headers.push(k)));
    ws.columns = headers.map((hd) => ({ header: hd, key: hd, width: Math.min(60, Math.max(12, hd.length + 2)) }));
    rows.forEach((r) => ws.addRow(r));
    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.eachRow((row, i) => i > 1 && (row.alignment = { wrapText: true, vertical: 'top' }));
  });
  const buffer = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer, filename: `${name}.xlsx`, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
}

module.exports = { build, KINDS };
