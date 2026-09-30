let excelJsPromise;

async function excelJs() {
  if (!globalThis.ExcelJS) {
    excelJsPromise ??= import('../../vendor/exceljs.min.js');
    await excelJsPromise;
  }
  if (!globalThis.ExcelJS?.Workbook) throw new Error('The local Excel writer could not be loaded.');
  return globalThis.ExcelJS;
}

async function workbookFileHandle() {
  const { dbGet } = await import('./db.js');
  const handle = await dbGet('handles', 'workbook');
  if (!handle) throw new Error('Choose an Excel workbook in Settings before capturing.');
  const permission = await handle.queryPermission({ mode: 'readwrite' });
  if (permission !== 'granted') {
    throw new Error('Workbook access is not currently granted. Open Settings and select the workbook again.');
  }
  return handle;
}

function toBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

const HEADERS = ['Screenshot', 'Timestamp', 'Page Title', 'Page URL', 'File Name', 'Capture ID'];

export async function appendCaptureToWorkbook(settings, item) {
  const handle = await workbookFileHandle();
  const file = await handle.getFile();
  const ExcelJS = await excelJs();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  const sheetName = settings.worksheet || 'Screenshots';
  const sheet = workbook.getWorksheet(sheetName) || workbook.addWorksheet(sheetName);
  const header = sheet.getRow(1);
  const columns = new Map();
  for (let col = 1; col <= Math.max(header.cellCount, HEADERS.length); col += 1) {
    const value = String(header.getCell(col).value ?? '').trim().toLowerCase();
    if (value) columns.set(value, col);
  }
  for (const name of HEADERS) {
    const key = name.toLowerCase();
    if (!columns.has(key)) {
      const col = Math.max(header.cellCount, ...columns.values(), 0) + 1;
      header.getCell(col).value = name;
      columns.set(key, col);
    }
  }

  const captureIdColumn = columns.get('capture id');
  sheet.getColumn(columns.get('screenshot')).width = 28;
  for (const [name, width] of [['timestamp', 22], ['page title', 36], ['page url', 48], ['file name', 32], ['capture id', 38]]) {
    sheet.getColumn(columns.get(name)).width = width;
  }
  for (let rowNum = 2; rowNum <= sheet.rowCount; rowNum += 1) {
    if (sheet.getRow(rowNum).getCell(captureIdColumn).value === item.id) {
      return { row: rowNum, duplicate: true };
    }
  }

  const values = [];
  const put = (name, value) => { values[columns.get(name.toLowerCase()) - 1] = value; };
  put('Screenshot', 'Embedded image');
  put('Timestamp', settings.includeTimestamp ? item.timestamp : '');
  put('Page Title', settings.includePageTitle ? item.pageTitle : '');
  put('Page URL', settings.includePageUrl ? item.pageUrl : '');
  put('File Name', item.fileName);
  put('Capture ID', item.id);
  const row = sheet.addRow(values);
  row.height = 78;

  const imageId = workbook.addImage({
    base64: toBase64(await item.blob.arrayBuffer()),
    extension: 'png'
  });
  sheet.addImage(imageId, {
    tl: { col: columns.get('screenshot') - 1, row: row.number - 1 },
    ext: { width: 180, height: 100 },
    editAs: 'oneCell'
  });

  const output = await workbook.xlsx.writeBuffer();
  const writable = await handle.createWritable();
  try {
    await writable.write(output);
    await writable.close();
  } catch (error) {
    await writable.abort().catch(() => {});
    throw error;
  }
  return { row: row.number, duplicate: false };
}
