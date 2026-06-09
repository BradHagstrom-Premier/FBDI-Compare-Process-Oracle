// Applaud snapshot extractor (Step A, programmatic).
//
// Reads an Applaud .mdb DIRECTLY with `mdb-reader` — the exact pure-JS library
// the applaud-mcp server uses — and dumps the raw rows needed to assemble a
// scoped snapshot to JSON. This replaces agent-driven (re-typed) extraction,
// which does not scale past a few hundred rows: the only path from an MCP query
// result to disk is the agent re-emitting it, so bulk volume blows the context
// budget AND can't be fidelity-checked against the source. Reading the file with
// the same library, programmatically, is byte-exact and reusable.
//
// Provenance note: this uses the SAME reader (mdb-reader), file, and password as
// applaud-mcp; it is the "headless extraction" path the audit design anticipated.
// As a cross-check, each object's row count is asserted against the COUNT(*)
// values measured live via applaud-mcp — if the two data paths agree, the read
// is faithful with zero transcription.
//
// Usage (from repo root):
//   node tools/extract_applaud_snapshot.mjs
// Requires the ApplaudMCP node_modules (for mdb-reader) at APPLAUD_MCP_DIR below.

import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const APPLAUD_MCP_DIR = 'C:/Users/10193/Definian/ApplaudMCP';
const MDB_PATH = 'C:/Users/10193/Definian/MDB_for_ApplaudMCP/ORACLE_MASTER/AP0STE.mdb';
const MDB_PASSWORD = 'sailboat';
const OUT_PATH = 'baselines/applaud/raw/extract.json';

// Resolve mdb-reader from ApplaudMCP's node_modules (it is not a dep of this repo).
const require = createRequire(APPLAUD_MCP_DIR + '/index.js');
const mdbReaderEntry = require.resolve('mdb-reader');
const { default: MDBReader } = await import(pathToFileURL(mdbReaderEntry).href);

// Pilot scope: 10 confirmed tables. table | prefix | import file | export file(s) | expected COUNT(*) (from applaud-mcp).
const TABLES = [
  { table: 'T_AP_INVOICE_INT',           prefix: 'TA1', if_: 'I_T_AP_INVOICE_INT',           efs: ['X_T_AP_INVOICE_INT', 'X_T_AP_INVOICE_INT_TXT'], exp: { db: 155, dd: 131, imp: 131, exp: { 'X_T_AP_INVOICE_INT': 131, 'X_T_AP_INVOICE_INT_TXT': 132 } } },
  { table: 'T_AP_INVOICE_LINES',         prefix: 'T99', if_: 'I_T_AP_INVOICE_LINES',         efs: ['X_T_AP_INVOICE_LINES'],                         exp: { db: 189, dd: 162, imp: 151, exp: { 'X_T_AP_INVOICE_LINES': 151 } } },
  { table: 'T_BANKS_BRANCHES',           prefix: 'T32', if_: 'I_T_BANKS_BRANCHES',           efs: ['T_BANKS_BRANCHES'],                             exp: { db: 49,  dd: 23,  imp: 23,  exp: { 'T_BANKS_BRANCHES': 23 } } },
  { table: 'T_BPA_PO_LINES_INTERFACE',   prefix: 'T64', if_: 'I_T_BPA_PO_LINES_INTERFACE',   efs: ['X_T_BPA_PO_LINES_INTERFACE'],                   exp: { db: 132, dd: 106, imp: 104, exp: { 'X_T_BPA_PO_LINES_INTERFACE': 104 } } },
  { table: 'T_EGP_COMPONENTS_INTERFACE', prefix: 'T91', if_: 'I_T_EGP_COMPONENTS_INTERFACE', efs: ['X_T_EGP_COMPONENTS_INTERFACE'],                 exp: { db: 126, dd: 98,  imp: 98,  exp: { 'X_T_EGP_COMPONENTS_INTERFACE': 98 } } },
  { table: 'T_EGP_ITEM_CATEGORIES_INT',  prefix: 'T87', if_: 'I_T_EGP_ITEM_CATEGORIES_INT',  efs: ['X_T_EGP_ITEM_CATEGORIES_INT'],                  exp: { db: 40,  dd: 14,  imp: 12,  exp: { 'X_T_EGP_ITEM_CATEGORIES_INT': 14 } } },
  { table: 'T_EGO_ITEM_INTF_EFF_B',      prefix: 'T86', if_: 'I_T_EGO_ITEM_INTF_EFF_B',      efs: ['X_T_EGO_ITEM_INTF_EFF_B'],                      exp: { db: 156, dd: 130, imp: 130, exp: { 'X_T_EGO_ITEM_INTF_EFF_B': 130 } } },
  { table: 'T_MSC_ST_ASSIGNMENT_SETS',   prefix: 'T04', if_: 'I_T_MSC_ST_ASSIGNMENT_SETS',   efs: ['X_T_MSC_ST_ASSIGNMENT_SETS'],                   exp: { db: 53,  dd: 27,  imp: 27,  exp: { 'X_T_MSC_ST_ASSIGNMENT_SETS': 27 } } },
  { table: 'T_POZ_SUPPLIERS_INT',        prefix: 'T07', if_: 'I_T_POZ_SUPPLIERS_INT',        efs: ['X_T_POZ_SUPPLIERS'],                            exp: { db: 184, dd: 156, imp: 155, exp: { 'X_T_POZ_SUPPLIERS': 156 } } },
  { table: 'T_POZ_SUPPLIER_SITES_INT',   prefix: 'T09', if_: 'I_T_POZ_SUPPLIER_SITES_INT',   efs: ['X_T_POZ_SUPPLIER_SITES'],                       exp: { db: 226, dd: 199, imp: 155, exp: { 'X_T_POZ_SUPPLIER_SITES': 199 } } },
];

const reader = new MDBReader(readFileSync(MDB_PATH), { password: MDB_PASSWORD });
const dbDetail = reader.getTable('DatabaseDetail').getData();
const dataDict = reader.getTable('DataDictionary').getData();
const impDetail = reader.getTable('ImportDetail').getData();
const expDetail = reader.getTable('ExportDetail').getData();

const out = { _meta: { mdb_path: MDB_PATH, system: 'ORACLE_MASTER' },
              database_detail: {}, data_dictionary: {}, import_detail: {}, export_detail: {} };

const mismatches = [];
function check(label, got, want) {
  const ok = got === want;
  if (!ok) mismatches.push(`${label}: read ${got} but applaud-mcp COUNT(*)=${want}`);
  console.log(`${ok ? 'OK ' : 'XX '} ${label.padEnd(48)} read=${String(got).padStart(4)}  count=${String(want).padStart(4)}`);
}

for (const t of TABLES) {
  const db = dbDetail.filter(r => r.Name === t.table)
    .map(r => ({ Row: r.Row, DDID: r.DDID, ODBCName: r.ODBCName }));
  out.database_detail[t.table] = db;
  check(`DBDetail ${t.table}`, db.length, t.exp.db);

  // LIKE '<prefix>%' equivalent: a 3-char TableId prefix; startsWith excludes '@'-audit DDIDs.
  const dd = dataDict.filter(r => typeof r.Name === 'string' && r.Name.startsWith(t.prefix))
    .map(r => ({ Name: r.Name, DataType: r.DataType, Size: r.Size, DecPlaces: r.DecPlaces }));
  out.data_dictionary[t.prefix] = dd;
  check(`DataDictionary ${t.prefix}%`, dd.length, t.exp.dd);

  const imp = impDetail.filter(r => r.Name === t.if_)
    .map(r => ({ Row: r.Row, DDID: r.DDID, Pic: r.Pic, InputType: r.InputType }));
  out.import_detail[t.if_] = imp;
  check(`ImportDetail ${t.if_}`, imp.length, t.exp.imp);

  for (const ef of t.efs) {
    const ex = expDetail.filter(r => r.Name === ef)
      .map(r => ({ Row: r.Row, DDID: r.DDID, Pic: r.Pic, ColumnHeader: r.ColumnHeader }));
    out.export_detail[ef] = ex;
    check(`ExportDetail ${ef}`, ex.length, t.exp.exp[ef]);
  }
}

mkdirSync('baselines/applaud/raw', { recursive: true });
writeFileSync(OUT_PATH, JSON.stringify(out, null, 0), 'utf-8');

console.log(`\nWrote ${OUT_PATH}`);
if (mismatches.length) {
  console.error(`\nCOUNT MISMATCHES (${mismatches.length}) — read path disagrees with applaud-mcp:`);
  for (const m of mismatches) console.error('  ' + m);
  process.exit(1);
}
console.log('\nAll object counts match applaud-mcp COUNT(*). Extraction faithful.');
