/**
 * DevShop Report Studio - Electron Main Process
 * Replaces Flask + subprocess architecture entirely.
 * All document generation runs in-process via docx library.
 */

const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs   = require('fs');
const os   = require('os');

// ── Output directory ────────────────────────────────────────────────────────
const OUTPUT_DIR = path.join(os.homedir(), 'Documents', 'DevShop Studio', 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

// ── Window ──────────────────────────────────────────────────────────────────
let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width:  1400,
    height: 900,
    minWidth:  900,
    minHeight: 600,
    frame: false,           // custom titlebar
    titleBarStyle: 'hidden',
    backgroundColor: '#0A0E17',
    webPreferences: {
      preload: path.join(__dirname, 'src', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    icon: path.join(__dirname, 'assets', 'icon.png')
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'index.html'));

  // Open DevTools only in dev mode
  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ── Window controls (custom titlebar) ───────────────────────────────────────
ipcMain.on('window-minimize', () => mainWindow.minimize());
ipcMain.on('window-maximize', () => {
  if (mainWindow.isMaximized()) mainWindow.unmaximize();
  else mainWindow.maximize();
});
ipcMain.on('window-close', () => mainWindow.close());

// ── Helper: escape strings for JS embedding ─────────────────────────────────
function esc(s) {
  if (s == null) return '';
  return String(s)
    .replace(/\\/g, '\\\\')
    .replace(/'/g,  "\\'")
    .replace(/\n/g, ' ')
    .replace(/\r/g, '');
}

function pcol(p) {
  return (p && (p.includes('1 - Critical') || p.includes('2 - High')))
    ? '991B1B' : '374151';
}

// ── Validate ─────────────────────────────────────────────────────────────────
ipcMain.handle('validate', async (_event, data) => {
  try {
    const required = ['meta', 'executive_summary', 'metrics', 'incidents', 'stories'];
    const missing  = required.filter(k => !(k in data));
    const warnings = [];

    if (!data.meta?.account)       warnings.push('Account name missing from meta');
    if (!data.meta?.week_of)       warnings.push('Week of date missing from meta');
    if (!data.focus_item?.title)   warnings.push('No focus item defined');
    if (!data.incidents?.length)   warnings.push('No incidents — section will show empty');
    if (!data.action_items?.length) warnings.push('No action items defined');

    return {
      valid:    missing.length === 0,
      missing,
      warnings,
      counts: {
        incidents:    (data.incidents    || []).length,
        stories:      (data.stories      || []).length,
        blockers:     (data.blockers     || []).length,
        action_items: (data.action_items || []).length,
        team:         (data.team         || []).length,
        uat_backlog:  (data.uat_backlog  || []).length
      }
    };
  } catch (e) {
    return { error: e.message };
  }
});

// ── Generate ─────────────────────────────────────────────────────────────────
ipcMain.handle('generate', async (_event, data) => {
  try {
    const {
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
      LevelFormat, Footer, Header, TabStopType
    } = require('docx');

    const meta      = data.meta || {};
    const account   = esc(meta.account || 'DevShop');
    const week_of   = esc(meta.week_of || new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }));
    const pm        = esc(meta.pm || 'Program Manager');
    const org       = esc(meta.org || '');
    const summary   = esc(data.executive_summary || '');

    const m         = data.metrics || {};
    const incidents = data.incidents    || [];
    const stories   = data.stories      || [];
    const s_prod    = stories.filter(s => s.phase === 'PROD');
    const s_test    = stories.filter(s => s.phase === 'TEST');
    const s_wip     = stories.filter(s => ['WIP','ENHC'].includes(s.phase));
    const uat       = data.uat_backlog   || [];
    const blockers  = data.blockers      || [];
    const actions   = data.action_items  || [];
    const team      = data.team          || [];
    const fi        = data.focus_item    || {};

    const footer_left  = org || account;
    const header_title = account + ' DevShop - Status Report';
    const cover_sub    = account + ' CSM Engagement   |   Week of ' + week_of;
    const cover_credit = 'Prepared by: ' + pm;
    const footer_txt   = footer_left + ' | Confidential';
    const closing_txt  = 'Prepared by ' + pm + '  |  ' + account + ' CSM DevShop  |  ' + week_of;

    // Colors
    const NAVY='1F3864', BLUE='2E5FAC', LIGHT_BLUE='DBEAFE', TEAL='1A6B6B',
          LIGHT_TEAL='E0F2F1', GREEN='166534', LIGHT_GREEN='DCFCE7',
          AMBER='92400E', LIGHT_AMBER='FEF3C7', RED='991B1B', LIGHT_RED='FEE2E2',
          GRAY='374151', LIGHT_GRAY='F3F4F6', MID_GRAY='D1D5DB', WHITE='FFFFFF',
          ORANGE='C2410C', LIGHT_ORANGE='FED7AA', PURPLE='5B21B6', LIGHT_PURPLE='EDE9FE';

    const bdr = { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY };
    const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
    const noBdr = { style: BorderStyle.NONE, size: 0, color: WHITE };
    const noBorders = { top: noBdr, bottom: noBdr, left: noBdr, right: noBdr };

    function hdrCell(t, bg, tc, w) {
      return new TableCell({
        borders, width: { size: w, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.LEFT,
          children: [new TextRun({ text: t, bold: true, color: tc, size: 18, font: 'Arial' })] })]
      });
    }

    function dataCell(t, bg, tc, w, bold) {
      return new TableCell({
        borders, width: { size: w, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: String(t || ''), bold: !!bold, color: tc, size: 18, font: 'Arial' })]
        })]
      });
    }

    function statusCell(t, w) {
      let bg, tc;
      if (t === 'In Progress' || t === 'Work in Progress') { bg = LIGHT_BLUE;   tc = BLUE;   }
      else if (t === 'Validate in TEST' || t === 'QA Testing') { bg = LIGHT_AMBER; tc = AMBER;  }
      else if (t.includes('PROD') || t === 'Move to PROD')   { bg = LIGHT_GREEN; tc = GREEN;  }
      else if (t === 'On Hold')  { bg = LIGHT_GRAY;   tc = GRAY;   }
      else if (t === 'Open')     { bg = LIGHT_ORANGE; tc = ORANGE; }
      else if (t === 'Draft')    { bg = LIGHT_PURPLE; tc = PURPLE; }
      else                       { bg = LIGHT_GRAY;   tc = GRAY;   }
      return dataCell(t, bg, tc, w, true);
    }

    function h2(t, color) {
      return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 },
        children: [new TextRun({ text: t, bold: true, color: color || BLUE, size: 28, font: 'Arial' })]
      });
    }

    function para(t, bold, color, size) {
      return new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: t, bold: !!bold, color: color || GRAY, size: size || 20, font: 'Arial' })]
      });
    }

    function bullet(t) {
      return new Paragraph({
        numbering: { reference: 'bullets', level: 0 },
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: t, color: GRAY, size: 19, font: 'Arial' })]
      });
    }

    function divider() {
      return new Paragraph({
        spacing: { before: 80, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_GRAY, space: 1 } },
        children: []
      });
    }

    function sp() {
      return new Paragraph({ spacing: { before: 160, after: 0 }, children: [] });
    }

    // Row builders
    function incRows() {
      if (!incidents.length) return [new TableRow({ children: [dataCell('No incidents', WHITE, GRAY, 9360)] })];
      return incidents.map((inc, i) => {
        const bg = i % 2 === 0 ? WHITE : LIGHT_GRAY;
        return new TableRow({ children: [
          dataCell(inc.num||'',      bg, BLUE,             1200, true),
          dataCell(inc.title||'',    bg, GRAY,             2400),
          dataCell(inc.assignee||'', bg, GRAY,             1000),
          statusCell(inc.state||'',                        1200),
          dataCell(inc.priority||'', bg, pcol(inc.priority), 800),
          dataCell(inc.note||'',     bg, GRAY,             2760)
        ]});
      });
    }

    function storyRows(items, alt) {
      if (!items.length) return [new TableRow({ children: [dataCell('None', WHITE, GRAY, 9360)] })];
      return items.map((s, i) => {
        const bg = i % 2 === 0 ? WHITE : alt;
        return new TableRow({ children: [
          dataCell(s.num||'',      bg, TEAL,            1300, true),
          dataCell(s.title||'',    bg, GRAY,            3500),
          dataCell(s.assignee||'', bg, GRAY,            1200),
          dataCell(s.priority||'', bg, pcol(s.priority),1200),
          statusCell(s.state||'',                       2160)
        ]});
      });
    }

    function blockerRows() {
      if (!blockers.length) return [new TableRow({ children: [dataCell('No blockers', WHITE, GRAY, 9360)] })];
      return blockers.map((b, i) => {
        const bg = i % 2 === 0 ? WHITE : LIGHT_GRAY;
        return new TableRow({ children: [
          dataCell(b.blocker||'',    bg,          GRAY,  3000),
          dataCell(b.impact||'',     LIGHT_AMBER, AMBER, 3180),
          dataCell(b.resolution||'', bg,          GRAY,  3180)
        ]});
      });
    }

    function actionRows() {
      if (!actions.length) return [new TableRow({ children: [dataCell('No action items', WHITE, GRAY, 9360)] })];
      return actions.map((a, i) => {
        const bg = i % 2 === 0 ? WHITE : LIGHT_GRAY;
        return new TableRow({ children: [
          dataCell(a.owner||'',  bg, GRAY, 2400, true),
          dataCell(a.action||'', bg, GRAY, 3960),
          dataCell(a.item||'',   bg, BLUE, 1500),
          dataCell(a.due||'',    bg, GRAY, 1500)
        ]});
      });
    }

    function teamRows() {
      if (!team.length) return [new TableRow({ children: [dataCell('No team data', WHITE, GRAY, 9360)] })];
      return team.map((t, i) => {
        const bg = i % 2 === 0 ? WHITE : LIGHT_GRAY;
        return new TableRow({ children: [
          dataCell(t.name||'',  bg, NAVY, 2340, true),
          dataCell(t.role||'',  bg, GRAY, 2340),
          dataCell(t.focus||'', bg, GRAY, 4680)
        ]});
      });
    }

    const fi_bullets_nodes = (fi.bullets || []).map(b => bullet(esc(b)));
    const uat_bullet_nodes = uat.map(u => bullet(esc(u)));

    const doc = new Document({
      numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] }] },
      styles: {
        default: { document: { run: { font: 'Arial', size: 20, color: GRAY } } },
        paragraphStyles: [{ id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, font: 'Arial', color: BLUE }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } }]
      },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
        headers: { default: new Header({ children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
            spacing: { before: 0, after: 120 },
            tabStops: [{ type: TabStopType.RIGHT, position: 8640 }],
            children: [
              new TextRun({ text: header_title, bold: true, color: NAVY, size: 18, font: 'Arial' }),
              new TextRun({ text: '\t' + week_of, color: GRAY, size: 18, font: 'Arial' })
            ]
          })
        ]}) },
        footers: { default: new Footer({ children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_GRAY, space: 1 } },
            spacing: { before: 80 },
            children: [new TextRun({ text: footer_txt, color: GRAY, size: 16, font: 'Arial' })]
          })
        ]}) },
        children: [
          // Cover
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
            new TableRow({ children: [new TableCell({
              borders: noBorders,
              shading: { fill: NAVY, type: ShadingType.CLEAR },
              width: { size: 9360, type: WidthType.DXA },
              margins: { top: 360, bottom: 360, left: 360, right: 360 },
              children: [
                new Paragraph({ spacing: { before: 0, after: 100 }, children: [new TextRun({ text: 'DEVSHOP STATUS REPORT', bold: true, color: WHITE, size: 48, font: 'Arial' })] }),
                new Paragraph({ spacing: { before: 0, after: 80  }, children: [new TextRun({ text: cover_sub, color: 'AECBF0', size: 22, font: 'Arial' })] }),
                new Paragraph({ children: [new TextRun({ text: cover_credit, color: 'C0D8F5', size: 20, font: 'Arial' })] })
              ]
            })] })
          ]}),

          sp(), h2('Executive Summary'),
          para(summary, false, GRAY, 19), sp(),

          // Metrics
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1560,1560,1560,1560,1560,1560], rows: [
            new TableRow({ children: [
              hdrCell('Total Tasks',    NAVY,  WHITE, 1560),
              hdrCell('In Progress',   BLUE,  WHITE, 1560),
              hdrCell('In Testing',    TEAL,  WHITE, 1560),
              hdrCell('Ready for PROD',GREEN, WHITE, 1560),
              hdrCell('On Hold',       GRAY,  WHITE, 1560),
              hdrCell('Critical Priority', RED, WHITE, 1560)
            ]}),
            new TableRow({ children: [
              dataCell(String(m.total_tasks      || 0), LIGHT_BLUE,  NAVY,  1560, true),
              dataCell(String(m.in_progress      || 0), LIGHT_BLUE,  BLUE,  1560, true),
              dataCell(String(m.in_testing       || 0), LIGHT_TEAL,  TEAL,  1560, true),
              dataCell(String(m.ready_for_prod   || 0), LIGHT_GREEN, GREEN, 1560, true),
              dataCell(String(m.on_hold          || 0), LIGHT_GRAY,  GRAY,  1560, true),
              dataCell(String(m.critical_priority|| 0), LIGHT_RED,   RED,   1560, true)
            ]})
          ]}),

          sp(), divider(), h2('Focus Item'),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
            new TableRow({ children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 8, color: AMBER }, bottom: bdr, left: { style: BorderStyle.SINGLE, size: 8, color: AMBER }, right: bdr },
              shading: { fill: LIGHT_AMBER, type: ShadingType.CLEAR },
              width: { size: 9360, type: WidthType.DXA },
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [
                new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: esc(fi.title||''), bold: true, color: AMBER, size: 22, font: 'Arial' })] }),
                new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'Status: ' + esc(fi.status||''), bold: true, color: ORANGE, size: 19, font: 'Arial' })] }),
                ...fi_bullets_nodes
              ]
            })] })
          ]}),

          sp(), divider(),
          h2('Active Incidents'),
          para(String(incidents.length) + ' incidents currently open.', false, GRAY, 19), sp(),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200,2400,1000,1200,800,2760], rows: [
            new TableRow({ children: [
              hdrCell('Number',NAVY,WHITE,1200), hdrCell('Description',NAVY,WHITE,2400),
              hdrCell('Assignee',NAVY,WHITE,1000), hdrCell('State',NAVY,WHITE,1200),
              hdrCell('Priority',NAVY,WHITE,800), hdrCell('Notes / Next Action',NAVY,WHITE,2760)
            ]}),
            ...incRows()
          ]}),

          sp(), divider(), h2('Stories Ready for Production', GREEN),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1300,3500,1200,1200,2160], rows: [
            new TableRow({ children: [hdrCell('Story #',GREEN,WHITE,1300),hdrCell('Description',GREEN,WHITE,3500),hdrCell('Assignee',GREEN,WHITE,1200),hdrCell('Priority',GREEN,WHITE,1200),hdrCell('State',GREEN,WHITE,2160)] }),
            ...storyRows(s_prod, LIGHT_GREEN)
          ]}),

          sp(), h2('Stories Currently in Testing', TEAL),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1300,3500,1200,1200,2160], rows: [
            new TableRow({ children: [hdrCell('Story #',TEAL,WHITE,1300),hdrCell('Description',TEAL,WHITE,3500),hdrCell('Assignee',TEAL,WHITE,1200),hdrCell('Priority',TEAL,WHITE,1200),hdrCell('State',TEAL,WHITE,2160)] }),
            ...storyRows(s_test, LIGHT_TEAL)
          ]}),

          sp(), h2('Stories in Development / Enhancements In-Flight', BLUE),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1300,3500,1200,1200,2160], rows: [
            new TableRow({ children: [hdrCell('Story #',BLUE,WHITE,1300),hdrCell('Description',BLUE,WHITE,3500),hdrCell('Assignee',BLUE,WHITE,1200),hdrCell('Priority',BLUE,WHITE,1200),hdrCell('State',BLUE,WHITE,2160)] }),
            ...storyRows(s_wip, LIGHT_BLUE)
          ]}),

          sp(), divider(), h2('UAT Pending - Stakeholder Action Required'),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
            new TableRow({ children: [new TableCell({
              borders, shading: { fill: LIGHT_AMBER, type: ShadingType.CLEAR },
              width: { size: 9360, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 200, right: 200 },
              children: [para('The following stories have completed internal testing but remain open pending UAT sign-off.', false, GRAY, 19)]
            })] })
          ]}),
          sp(), ...uat_bullet_nodes,

          sp(), divider(), h2('Process Update'),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
            new TableRow({ children: [new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 8, color: BLUE }, bottom: bdr, left: bdr, right: bdr },
              shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
              width: { size: 9360, type: WidthType.DXA }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [para(esc(data.process_note || ''), false, GRAY, 19)]
            })] })
          ]}),

          sp(), divider(), h2('Blockers and Risks'),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3000,3180,3180], rows: [
            new TableRow({ children: [hdrCell('Blocker / Risk',RED,WHITE,3000),hdrCell('Impact',RED,WHITE,3180),hdrCell('Resolution Path',RED,WHITE,3180)] }),
            ...blockerRows()
          ]}),

          sp(), divider(), h2('Action Items'),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2400,3960,1500,1500], rows: [
            new TableRow({ children: [hdrCell('Owner',NAVY,WHITE,2400),hdrCell('Action',NAVY,WHITE,3960),hdrCell('Related Item',NAVY,WHITE,1500),hdrCell('Due / Status',NAVY,WHITE,1500)] }),
            ...actionRows()
          ]}),

          sp(), divider(), h2('Team Overview'),
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340,2340,4680], rows: [
            new TableRow({ children: [hdrCell('Name',NAVY,WHITE,2340),hdrCell('Role',NAVY,WHITE,2340),hdrCell('Current Focus',NAVY,WHITE,4680)] }),
            ...teamRows()
          ]}),

          sp(),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: closing_txt, color: MID_GRAY, size: 17, font: 'Arial', italics: true })]
          })
        ]
      }]
    });

    // Write file
    const date_str  = new Date().toLocaleDateString('en-US', { month:'short', day:'2-digit', year:'numeric' }).replace(/ /g,'').replace(',','_');
    const safe_acct = (meta.account || 'DevShop').replace(/\s+/g, '_');
    const filename  = safe_acct + '_Status_Report_' + date_str + '.docx';
    const outPath   = path.join(OUTPUT_DIR, filename);

    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buf);

    return { success: true, filename, path: outPath };
  } catch (e) {
    return { error: e.message };
  }
});

// ── Open output file in system app ───────────────────────────────────────────
ipcMain.handle('open-file', async (_event, filePath) => {
  await shell.openPath(filePath);
  return true;
});

// ── Open output folder ───────────────────────────────────────────────────────
ipcMain.handle('open-output-folder', async () => {
  await shell.openPath(OUTPUT_DIR);
  return true;
});

// ── Save file dialog (download) ──────────────────────────────────────────────
ipcMain.handle('save-file', async (_event, srcPath, defaultName) => {
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title:       'Save Report',
    defaultPath: path.join(os.homedir(), 'Documents', defaultName),
    filters:     [{ name: 'Word Documents', extensions: ['docx'] }]
  });
  if (canceled || !filePath) return { canceled: true };
  fs.copyFileSync(srcPath, filePath);
  return { success: true, path: filePath };
});

// ── Generate Call Script ──────────────────────────────────────────────────────
ipcMain.handle('generate-script', async (_event, data) => {
  try {
    const {
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
      Footer, Header, TabStopType
    } = require('docx');

    const meta      = data.meta || {};
    const account   = esc(meta.account   || 'Account');
    const call_type = esc(meta.call_type || 'Client Update Call');
    const call_date = esc(meta.date      || '');
    const pm        = esc(meta.prepared_by || meta.pm || 'Program Manager');
    const audience  = esc(meta.audience  || '');
    const presenter = esc(meta.presenter_role || '');
    const topics    = data.topics || [];

    const NAVY='1F3864', BLUE='2E5FAC', LIGHT_BLUE='DBEAFE', TEAL='1A6B6B',
          LIGHT_TEAL='E0F2F1', GREEN='166534', LIGHT_GREEN='DCFCE7',
          AMBER='92400E', LIGHT_AMBER='FEF3C7', RED='991B1B',
          GRAY='374151', LIGHT_GRAY='F3F4F6', MID_GRAY='D1D5DB', WHITE='FFFFFF',
          ORANGE='C2410C', LIGHT_ORANGE='FED7AA', PURPLE='5B21B6', LIGHT_PURPLE='EDE9FE';

    const bdr     = { style: BorderStyle.SINGLE, size: 1, color: MID_GRAY };
    const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
    const noBdr   = { style: BorderStyle.NONE, size: 0, color: WHITE };
    const noBorders = { top: noBdr, bottom: noBdr, left: noBdr, right: noBdr };

    function hdrCell(t, bg, tc, w) {
      return new TableCell({
        borders, width: { size: w, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, color: tc, size: 18, font: 'Arial' })] })]
      });
    }

    function dataCell(t, bg, tc, w, bold) {
      return new TableCell({
        borders, width: { size: w, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: String(t || ''), bold: !!bold, color: tc, size: 18, font: 'Arial' })] })]
      });
    }

    function statusCell(t, w) {
      let bg, tc;
      if (t === 'In Progress' || t === 'Work in Progress') { bg = LIGHT_BLUE;   tc = BLUE;   }
      else if (t === 'Validate in TEST' || t === 'QA Testing') { bg = LIGHT_AMBER; tc = AMBER;  }
      else if (t.includes('PROD') || t === 'Move to PROD')   { bg = LIGHT_GREEN; tc = GREEN;  }
      else if (t === 'On Hold')  { bg = LIGHT_GRAY;   tc = GRAY;   }
      else if (t === 'Open')     { bg = LIGHT_ORANGE; tc = ORANGE; }
      else if (t === 'Draft')    { bg = LIGHT_PURPLE; tc = PURPLE; }
      else                       { bg = LIGHT_GRAY;   tc = GRAY;   }
      return dataCell(t, bg, tc, w, true);
    }

    function h2(t, color) {
      return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 },
        children: [new TextRun({ text: t, bold: true, color: color || BLUE, size: 26, font: 'Arial' })]
      });
    }

    function para(t, bold, color, size) {
      return new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: t, bold: !!bold, color: color || GRAY, size: size || 20, font: 'Arial' })]
      });
    }

    function divider() {
      return new Paragraph({
        spacing: { before: 80, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_GRAY, space: 1 } },
        children: []
      });
    }

    function sp() { return new Paragraph({ spacing: { before: 160, after: 0 }, children: [] }); }

    // Build topic children nodes
    function buildTopicNodes(t, idx) {
      const nodes = [];
      const t_title = esc(t.title || ('Topic ' + (idx + 1)));
      const t_say   = esc(t.say_this || '');
      const t_notes = esc(t.technical_notes || '');
      const qa      = t.qa_pairs  || [];
      const tasks   = t.tasks     || [];

      nodes.push(h2('TOPIC ' + (idx + 1) + ' — ' + t_title, NAVY));
      nodes.push(sp());

      if (t_say) {
        nodes.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
          new TableRow({ children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 8, color: GREEN }, bottom: bdr, left: { style: BorderStyle.SINGLE, size: 8, color: GREEN }, right: bdr },
            shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
            width: { size: 9360, type: WidthType.DXA }, margins: { top: 160, bottom: 160, left: 220, right: 220 },
            children: [
              new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'SAY THIS:', bold: true, color: GREEN, size: 18, font: 'Arial' })] }),
              para(t_say, false, GRAY, 19)
            ]
          })] })
        ]}));
        nodes.push(sp());
      }

      if (tasks.length) {
        nodes.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1300,1100,1500,3500,1960], rows: [
          new TableRow({ children: [
            hdrCell('Task #', BLUE, WHITE, 1300), hdrCell('Type', BLUE, WHITE, 1100),
            hdrCell('Assigned To', BLUE, WHITE, 1500), hdrCell('Description', BLUE, WHITE, 3500),
            hdrCell('State', BLUE, WHITE, 1960)
          ]}),
          ...tasks.map((tk, i) => {
            const bg = i % 2 === 0 ? WHITE : LIGHT_GRAY;
            return new TableRow({ children: [
              dataCell(tk.num||'',      bg, BLUE,  1300, true),
              dataCell(tk.type||'',     bg, GRAY,  1100),
              dataCell(tk.assignee||'', bg, GRAY,  1500),
              dataCell(tk.title||'',    bg, GRAY,  3500),
              statusCell(tk.state||'',        1960)
            ]});
          })
        ]}));
        nodes.push(sp());
      }

      if (qa.length) {
        nodes.push(new Paragraph({ spacing: { before: 120, after: 80 }, children: [new TextRun({ text: 'IF THEY ASK:', bold: true, color: AMBER, size: 19, font: 'Arial' })] }));
        nodes.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4500, 4860], rows: [
          new TableRow({ children: [hdrCell('Question', AMBER, WHITE, 4500), hdrCell('Quick Answer', AMBER, WHITE, 4860)] }),
          ...qa.map((q, i) => {
            const bg = i % 2 === 0 ? WHITE : 'F9FAFB';
            return new TableRow({ children: [
              dataCell(esc(q.question||''), bg, NAVY, 4500, true),
              dataCell(esc(q.answer||''),   bg, GRAY, 4860)
            ]});
          })
        ]}));
        nodes.push(sp());
      }

      if (t_notes) {
        nodes.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
          new TableRow({ children: [new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 6, color: GRAY }, bottom: bdr, left: { style: BorderStyle.SINGLE, size: 6, color: GRAY }, right: bdr },
            shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
            width: { size: 9360, type: WidthType.DXA }, margins: { top: 140, bottom: 140, left: 220, right: 220 },
            children: [
              new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'TECHNICAL NOTES (FOR YOUR EYES ONLY):', bold: true, color: GRAY, size: 16, font: 'Arial' })] }),
              para(t_notes, false, '6E7681', 18)
            ]
          })] })
        ]}));
        nodes.push(sp());
        nodes.push(divider());
      } else {
        nodes.push(divider());
      }

      return nodes;
    }

    const closing_txt = esc(data.closing_statement || '');
    const header_title = account + ' - ' + call_type;
    const footer_txt   = account + ' | Call Script | Confidential';

    const allTopicNodes = topics.flatMap((t, i) => buildTopicNodes(t, i));

    const closingNodes = closing_txt ? [
      sp(),
      new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
        new TableRow({ children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 8, color: TEAL }, bottom: bdr, left: { style: BorderStyle.SINGLE, size: 8, color: TEAL }, right: bdr },
          shading: { fill: LIGHT_TEAL, type: ShadingType.CLEAR },
          width: { size: 9360, type: WidthType.DXA }, margins: { top: 160, bottom: 160, left: 220, right: 220 },
          children: [
            new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'CLOSING:', bold: true, color: TEAL, size: 18, font: 'Arial' })] }),
            para(closing_txt, false, GRAY, 19)
          ]
        })] })
      ]})
    ] : [];

    const doc = new Document({
      styles: { default: { document: { run: { font: 'Arial', size: 20, color: GRAY } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
        headers: { default: new Header({ children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
            spacing: { before: 0, after: 120 },
            tabStops: [{ type: TabStopType.RIGHT, position: 8640 }],
            children: [
              new TextRun({ text: header_title, bold: true, color: NAVY, size: 18, font: 'Arial' }),
              new TextRun({ text: '\t' + call_date, color: GRAY, size: 18, font: 'Arial' })
            ]
          })
        ]}) },
        footers: { default: new Footer({ children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_GRAY, space: 1 } },
            spacing: { before: 80 },
            children: [new TextRun({ text: footer_txt, color: GRAY, size: 16, font: 'Arial' })]
          })
        ]}) },
        children: [
          // Cover
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows: [
            new TableRow({ children: [new TableCell({
              borders: noBorders, shading: { fill: NAVY, type: ShadingType.CLEAR },
              width: { size: 9360, type: WidthType.DXA }, margins: { top: 360, bottom: 360, left: 360, right: 360 },
              children: [
                new Paragraph({ spacing: { before: 0, after: 100 }, children: [new TextRun({ text: 'CLIENT UPDATE CALL SCRIPT', bold: true, color: WHITE, size: 48, font: 'Arial' })] }),
                new Paragraph({ spacing: { before: 0, after: 80  }, children: [new TextRun({ text: call_date + '   |   ' + audience, color: 'AECBF0', size: 22, font: 'Arial' })] }),
                new Paragraph({ children: [new TextRun({ text: 'Prepared by: ' + pm, color: 'C0D8F5', size: 20, font: 'Arial' })] })
              ]
            })] })
          ]}),
          sp(),
          // Call info table
          new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1800, 7560], rows: [
            new TableRow({ children: [hdrCell('Field', NAVY, WHITE, 1800), hdrCell('Detail', NAVY, WHITE, 7560)] }),
            new TableRow({ children: [dataCell('Client',         'F8FAFC', NAVY, 1800, true), dataCell(account,    'F8FAFC', GRAY, 7560)] }),
            new TableRow({ children: [dataCell('Call Type',      'FFFFFF', NAVY, 1800, true), dataCell(call_type,  'FFFFFF', GRAY, 7560)] }),
            new TableRow({ children: [dataCell('Date',           'F8FAFC', NAVY, 1800, true), dataCell(call_date,  'F8FAFC', GRAY, 7560)] }),
            new TableRow({ children: [dataCell('Prepared By',    'FFFFFF', NAVY, 1800, true), dataCell(pm,         'FFFFFF', GRAY, 7560)] }),
            new TableRow({ children: [dataCell('Presenter Role', 'F8FAFC', NAVY, 1800, true), dataCell(presenter,  'F8FAFC', GRAY, 7560)] }),
            new TableRow({ children: [dataCell('Audience',       'FFFFFF', NAVY, 1800, true), dataCell(audience,   'FFFFFF', GRAY, 7560)] })
          ]}),
          sp(),
          new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'HOW TO USE THIS SCRIPT', bold: true, color: NAVY, size: 20, font: 'Arial' })] }),
          para('This document is organized by topic. Each block contains the script language to speak out loud, questions you are likely to hear and how to answer them, and technical context for if things go deeper.', false, GRAY, 18),
          sp(), divider(),
          ...allTopicNodes,
          ...closingNodes,
          sp(),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: header_title + '  |  ' + call_date + '  |  Prepared by ' + pm, color: MID_GRAY, size: 16, font: 'Arial', italics: true })] })
        ]
      }]
    });

    const date_str  = new Date().toLocaleDateString('en-US', { month:'short', day:'2-digit', year:'numeric' }).replace(/ /g,'').replace(',','_');
    const safe_acct = (meta.account || 'Account').replace(/\s+/g, '_');
    const filename  = safe_acct + '_Call_Script_' + date_str + '.docx';
    const outPath   = path.join(OUTPUT_DIR, filename);

    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buf);

    return { success: true, filename, path: outPath };
  } catch (e) {
    return { error: e.message };
  }
});

// ── Generate RCA ──────────────────────────────────────────────────────────────
ipcMain.handle('generate-rca', async (_event, data) => {
  try {
    const {
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
      Footer, Header, TabStopType, LevelFormat, PageNumber
    } = require('docx');

    // ── Palette ────────────────────────────────────────────────
    const NAVY     = '1F3864';
    const BLUE     = '2E5FAC';
    const BLUE_LT  = 'DBEAFE';
    const GRAY     = '374151';
    const LGRAY    = 'F3F4F6';
    const MGRAY    = 'D1D5DB';
    const WHITE    = 'FFFFFF';
    const AMB_BG   = 'FFFBEB';
    const AMB_BD   = 'F59E0B';
    const AMB_TXT  = '92400E';
    const PH_TXT   = '94A3B8';
    const FONT     = 'Arial';
    const CW       = 9360; // content width DXA

    // ── Border helpers ─────────────────────────────────────────
    const thinBdr = { style: BorderStyle.SINGLE, size: 4, color: MGRAY };
    const cellBdr = { top: thinBdr, bottom: thinBdr, left: thinBdr, right: thinBdr };
    const noBdr   = { style: BorderStyle.NONE, size: 0, color: WHITE };
    const noBdrs  = { top: noBdr, bottom: noBdr, left: noBdr, right: noBdr };

    // ── Core helpers ───────────────────────────────────────────
    function sp(before, after) {
      return new Paragraph({ spacing: { before: before||0, after: after||0 }, children: [] });
    }

    function hrule(color) {
      return new Paragraph({
        spacing: { before: 40, after: 40 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: color||BLUE, space: 1 } },
        children: []
      });
    }

    // Numbered section heading — "1. Incident Overview"
    function secHead(num, title) {
      return new Paragraph({
        spacing: { before: 320, after: 120 },
        children: [
          new TextRun({ text: num + '.  ', font: FONT, size: 26, bold: true, color: NAVY }),
          new TextRun({ text: title,        font: FONT, size: 26, bold: true, color: BLUE })
        ]
      });
    }

    // Bold inline sub-label e.g. "Initial approach — ACL-based fix:"
    function subLabel(text) {
      return new Paragraph({
        spacing: { before: 180, after: 80 },
        children: [new TextRun({ text: text, font: FONT, size: 21, bold: true, color: NAVY })]
      });
    }

    // Body paragraph
    function bodyPara(text) {
      if (!text || !text.trim()) return null;
      return new Paragraph({
        spacing: { before: 80, after: 120 },
        children: [new TextRun({ text: text.trim(), font: FONT, size: 20, color: GRAY })]
      });
    }

    // Placeholder note when section is empty
    function placeholder(msg) {
      return new Paragraph({
        spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: '\u2014  ' + msg, font: FONT, size: 19, italics: true, color: PH_TXT })]
      });
    }

    // Callout box — amber bordered table
    function calloutBox(devName, items) {
      if (!items || !items.length) return null;
      return new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: [CW],
        rows: [new TableRow({ children: [new TableCell({
          width: { size: CW, type: WidthType.DXA },
          shading: { fill: AMB_BG, type: ShadingType.CLEAR },
          borders: {
            top:    { style: BorderStyle.SINGLE, size: 12, color: AMB_BD },
            bottom: { style: BorderStyle.SINGLE, size: 6,  color: AMB_BD },
            left:   { style: BorderStyle.SINGLE, size: 12, color: AMB_BD },
            right:  { style: BorderStyle.SINGLE, size: 6,  color: AMB_BD }
          },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({ spacing: { before: 0, after: 120 }, children: [
              new TextRun({ text: 'ACTION REQUIRED', font: FONT, size: 20, bold: true, color: AMB_BD }),
              new TextRun({ text: devName ? '  \u2014  ' + devName : '', font: FONT, size: 20, bold: true, color: AMB_TXT })
            ]}),
            ...items.map(item =>
              new Paragraph({
                numbering: { reference: 'callout-bullets', level: 0 },
                spacing: { before: 60, after: 60 },
                children: [new TextRun({ text: String(item||''), font: FONT, size: 19, color: AMB_TXT })]
              })
            )
          ]
        })]})],
      });
    }

    // Summary table row — label cell (navy) + value cell
    function sumRow(label, value, shade) {
      return new TableRow({ children: [
        new TableCell({
          width: { size: 2520, type: WidthType.DXA },
          shading: { fill: shade||NAVY, type: ShadingType.CLEAR },
          borders: cellBdr,
          margins: { top: 100, bottom: 100, left: 160, right: 160 },
          children: [new Paragraph({ children: [
            new TextRun({ text: String(label||''), font: FONT, size: 19, bold: true, color: WHITE })
          ]})]
        }),
        new TableCell({
          width: { size: 6840, type: WidthType.DXA },
          shading: { fill: 'F8FAFC', type: ShadingType.CLEAR },
          borders: cellBdr,
          margins: { top: 100, bottom: 100, left: 160, right: 160 },
          children: [new Paragraph({ children: [
            new TextRun({ text: String(value||'\u2014'), font: FONT, size: 19, color: GRAY })
          ]})]
        })
      ]});
    }

    // Timeline cells
    function tlHdr(text, w) {
      return new TableCell({
        width: { size: w, type: WidthType.DXA },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        borders: cellBdr,
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: [new Paragraph({ children: [
          new TextRun({ text: String(text||''), font: FONT, size: 19, bold: true, color: WHITE })
        ]})]
      });
    }

    function tlCell(text, w, shade) {
      return new TableCell({
        width: { size: w, type: WidthType.DXA },
        shading: { fill: shade||WHITE, type: ShadingType.CLEAR },
        borders: cellBdr,
        margins: { top: 90, bottom: 90, left: 160, right: 160 },
        children: [new Paragraph({ children: [
          new TextRun({ text: String(text||''), font: FONT, size: 18, color: GRAY })
        ]})]
      });
    }

    // ── Data ───────────────────────────────────────────────────
    const account  = String(data.account  || 'Client');
    const incident = String(data.incident || '');
    const title    = String(data.title    || incident);
    const pm       = 'Justin Alemazkour, Program Manager - Infocenter';
    const docDate  = new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'});

    // Build callout lookup by section key
    const coBySection = {};
    const coLoose     = [];
    (data.callouts || []).forEach(cg => {
      if (cg.section) {
        (coBySection[cg.section] = coBySection[cg.section] || []).push(cg);
      } else {
        coLoose.push(cg);
      }
    });

    function secCallouts(key) {
      return (coBySection[key] || []).map(cg => calloutBox(cg.dev, cg.items)).filter(Boolean);
    }

    // ── Document children ──────────────────────────────────────
    const ch = [];

    // ── Cover title block ──────────────────────────────────────
    ch.push(new Paragraph({
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'ROOT CAUSE ANALYSIS', font: FONT, size: 56, bold: true, color: NAVY })]
    }));
    ch.push(new Paragraph({
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: incident + (title && title !== incident ? ' \u2014 ' + title : ''), font: FONT, size: 24, color: BLUE })]
    }));
    ch.push(hrule(NAVY));
    ch.push(sp(160, 0));

    // ── Incident summary table ─────────────────────────────────
    ch.push(new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: [2520, 6840],
      rows: [
        sumRow('DOCUMENT',         'Root Cause Analysis'),
        sumRow('Incident',         data.incident    || ''),
        sumRow('Related Story',    data.story       || ''),
        sumRow('Change Record',    data.change      || ''),
        sumRow('Current Status',   data.status      || ''),
        sumRow('Assigned To',      data.developer   || ''),
        sumRow('Reported By',      data.reported_by || ''),
        sumRow('Date of Fix',      data.fix_date    || ''),
        sumRow('Prepared By',      pm),
        sumRow('Document Date',    docDate),
      ]
    }));

    ch.push(sp(240, 0));

    // ── Helper: render a narrative section ─────────────────────
    // Supports multi-block text (array) or single string
    // Each array item can be: string | { label, text } | { bullet: string } | { numbered: string }
    function renderSection(num, sectionTitle, content, sectionKey) {
      ch.push(secHead(num, sectionTitle));

      if (!content || (Array.isArray(content) && !content.length) ||
          (typeof content === 'string' && !content.trim())) {
        // Section is empty — show placeholder
        const placeholders = {
          's1': 'No incident overview provided. Describe what happened, when it was discovered, who was involved, and the initial impact observed.',
          's2': 'No root cause documented. Describe the underlying technical chain — configuration, code, system properties, or process gaps — and why the issue was not caught earlier.',
          's3': 'No impact assessment provided. Describe who or what was affected, the scope, data or access exposed, and the duration.',
          's4': 'No resolution details provided. Describe what was attempted, what the final fix was, and who developed, tested, and approved it.',
          's5': 'No contributing factors documented. Describe process gaps, missing controls, or environmental conditions that allowed this to occur or persist.',
          's6': 'No immediate actions recorded.',
          's7': 'No prevention plan documented.',
          's8': 'No open items recorded.',
        };
        ch.push(placeholder(placeholders[sectionKey] || 'Not yet documented.'));
      } else if (typeof content === 'string') {
        ch.push(bodyPara(content));
      } else if (Array.isArray(content)) {
        content.forEach(block => {
          if (!block) return;
          if (typeof block === 'string') {
            ch.push(bodyPara(block));
          } else if (block.label) {
            ch.push(subLabel(block.label));
            if (block.text) ch.push(bodyPara(block.text));
          } else if (block.bullet !== undefined) {
            ch.push(new Paragraph({
              numbering: { reference: 'doc-bullets', level: 0 },
              spacing: { before: 60, after: 60 },
              children: [new TextRun({ text: String(block.bullet||''), font: FONT, size: 20, color: GRAY })]
            }));
          } else if (block.numbered !== undefined) {
            ch.push(new Paragraph({
              numbering: { reference: 'doc-numbers', level: 0 },
              spacing: { before: 60, after: 60 },
              children: [new TextRun({ text: String(block.numbered||''), font: FONT, size: 20, color: GRAY })]
            }));
          }
        });
      }

      // Inline callouts for this section
      secCallouts(sectionKey).forEach(box => { ch.push(sp(120, 0)); ch.push(box); });
      ch.push(sp(80, 0));
    }

    // ── Sections 1–5: narrative ────────────────────────────────
    renderSection(1, 'Incident Overview',    data.s1, 's1');
    renderSection(2, 'Technical Root Cause', data.s2, 's2');
    renderSection(3, 'Impact Assessment',    data.s3, 's3');
    renderSection(4, 'Resolution Applied',   data.s4, 's4');
    renderSection(5, 'Contributing Factors', data.s5, 's5');

    // ── Section 6: Immediate Actions ──────────────────────────
    ch.push(secHead(6, 'Immediate Actions Taken'));
    if (data.s6 && data.s6.length) {
      data.s6.forEach(item => ch.push(
        new Paragraph({
          numbering: { reference: 'doc-bullets', level: 0 },
          spacing: { before: 60, after: 60 },
          children: [new TextRun({ text: String(item||''), font: FONT, size: 20, color: GRAY })]
        })
      ));
    } else {
      ch.push(placeholder('No immediate actions recorded.'));
    }
    secCallouts('s6').forEach(box => { ch.push(sp(120,0)); ch.push(box); });
    ch.push(sp(80,0));

    // ── Section 7: Prevention Plan ─────────────────────────────
    ch.push(secHead(7, 'Prevention Plan'));
    if (data.s7 && data.s7.length) {
      data.s7.forEach((item, i) => ch.push(
        new Paragraph({
          numbering: { reference: 'doc-numbers', level: 0 },
          spacing: { before: 60, after: 60 },
          children: [new TextRun({ text: String(item||''), font: FONT, size: 20, color: GRAY })]
        })
      ));
    } else {
      ch.push(placeholder('No prevention plan documented.'));
    }
    secCallouts('s7').forEach(box => { ch.push(sp(120,0)); ch.push(box); });
    ch.push(sp(80,0));

    // ── Section 8: Open Items ──────────────────────────────────
    ch.push(secHead(8, 'Open Items & Follow-Ups'));
    if (data.s8 && data.s8.length) {
      data.s8.forEach((item, i) => ch.push(
        new Paragraph({
          numbering: { reference: 'doc-numbers', level: 0 },
          spacing: { before: 60, after: 60 },
          children: [new TextRun({ text: String(item||''), font: FONT, size: 20, color: GRAY })]
        })
      ));
    } else {
      ch.push(placeholder('No open items recorded. If all follow-ups are complete, note that here.'));
    }
    secCallouts('s8').forEach(box => { ch.push(sp(120,0)); ch.push(box); });
    ch.push(sp(80,0));

    // ── Section 9 (numbered as 8 in example) Timeline ─────────
    ch.push(secHead(9, 'Incident Timeline'));
    if (data.timeline && data.timeline.length) {
      ch.push(new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: [2200, 7160],
        rows: [
          new TableRow({ children: [tlHdr('Time', 2200), tlHdr('Event', 7160)] }),
          ...data.timeline.map((t, i) => {
            const shade = i % 2 === 0 ? WHITE : 'F8FAFC';
            return new TableRow({ children: [
              tlCell(t.time  || '', 2200, shade),
              tlCell(t.event || '', 7160, shade)
            ]});
          })
        ]
      }));
    } else {
      ch.push(placeholder('No timeline entries provided. Add timestamps and events from the incident record.'));
    }
    ch.push(sp(80,0));

    // ── Loose callouts ─────────────────────────────────────────
    coLoose.forEach(cg => {
      const box = calloutBox(cg.dev, cg.items);
      if (box) { ch.push(sp(80,0)); ch.push(box); }
    });

    // ── Document notes + footer classification ─────────────────
    if (data.s10 && data.s10.trim()) {
      ch.push(hrule(MGRAY));
      ch.push(sp(80,0));
      ch.push(new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: 'Document Notes', font: FONT, size: 20, bold: true, color: GRAY })]
      }));
      ch.push(bodyPara(data.s10));
    }

    // Amber note at end
    ch.push(sp(160,0));
    ch.push(new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: [CW],
      rows: [new TableRow({ children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { fill: AMB_BG, type: ShadingType.CLEAR },
        borders: {
          top:    { style: BorderStyle.SINGLE, size: 8, color: AMB_BD },
          bottom: { style: BorderStyle.SINGLE, size: 8, color: AMB_BD },
          left:   { style: BorderStyle.SINGLE, size: 8, color: AMB_BD },
          right:  { style: BorderStyle.SINGLE, size: 8, color: AMB_BD }
        },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: [new Paragraph({ children: [
          new TextRun({ text: '\u26A0  All amber callout boxes must be resolved or removed before this document is distributed to the client.', font: FONT, size: 19, color: AMB_TXT })
        ]})]
      })]})],
    }));

    // ── Assemble document ──────────────────────────────────────
    const doc = new Document({
      numbering: {
        config: [
          {
            reference: 'doc-bullets',
            levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022',
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 },
                spacing: { before: 60, after: 60 } } } }]
          },
          {
            reference: 'doc-numbers',
            levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 },
                spacing: { before: 60, after: 60 } } } }]
          },
          {
            reference: 'callout-bullets',
            levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022',
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 360, hanging: 240 },
                spacing: { before: 60, after: 60 } } } }]
          }
        ]
      },
      styles: {
        default: { document: { run: { font: FONT, size: 20, color: GRAY } } },
      },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
          }
        },
        headers: { default: new Header({ children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
            spacing: { before: 0, after: 100 },
            tabStops: [{ type: TabStopType.RIGHT, position: CW }],
            children: [
              new TextRun({ text: 'CONFIDENTIAL  |  Root Cause Analysis  |  ' + account, font: FONT, size: 17, bold: true, color: NAVY }),
              new TextRun({ text: '\t' + incident, font: FONT, size: 17, color: GRAY })
            ]
          })
        ]}) },
        footers: { default: new Footer({ children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: MGRAY, space: 1 } },
            spacing: { before: 80 },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: 'Classification: Confidential. For internal and authorized client stakeholder use only.', font: FONT, size: 16, italics: true, color: GRAY })
            ]
          })
        ]}) },
        children: ch
      }]
    });

    const safe_acct = (data.account || 'Client').replace(/\s+/g, '_');
    const safe_inc  = (data.incident || 'RCA').replace(/[^a-zA-Z0-9_]/g, '');
    const filename  = 'RCA_' + safe_inc + '_' + safe_acct + '.docx';
    const outPath   = path.join(OUTPUT_DIR, filename);

    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buf);

    return { success: true, filename, path: outPath };
  } catch (e) {
    return { error: e.message };
  }
});
// ── List output files for History tab ────────────────────────────────────────
ipcMain.handle('list-output-files', async () => {
  try {
    if (!fs.existsSync(OUTPUT_DIR)) return [];
    const entries = fs.readdirSync(OUTPUT_DIR);
    return entries
      .filter(f => f.endsWith('.docx'))
      .map(f => {
        const full = path.join(OUTPUT_DIR, f);
        const stat = fs.statSync(full);
        return { name: f, path: full, modified: stat.mtimeMs, size: stat.size };
      })
      .sort((a, b) => b.modified - a.modified)
      .slice(0, 50); // max 50 entries
  } catch (e) {
    return [];
  }
});
