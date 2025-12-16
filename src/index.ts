
import { Hono } from 'hono';
import { serveStatic } from 'hono/cloudflare-pages';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

const app = new Hono();

// 1. 【重要修改】先定义 API 路由，防止被静态文件中间件拦截
app.post('/process', async (c) => {
  try {
    const body = await c.req.parseBody();
    
    // 获取上传的文件
    let targetFiles = body['targets'];
    const replacementFile = body['replacement'];
    const mode = body['mode'] as string;

    if (!targetFiles || !replacementFile) {
      return c.json({ error: 'Missing files' }, 400);
    }

    if (!Array.isArray(targetFiles)) {
      targetFiles = [targetFiles];
    }

    // @ts-ignore
    const repBuffer = await (replacementFile as File).arrayBuffer();
    const repWb = XLSX.read(repBuffer, { type: 'array' });
    const repSheet = repWb.Sheets[repWb.SheetNames[0]];
    const repData = XLSX.utils.sheet_to_json(repSheet, { header: ['key', 'value'], range: 1 });
    
    const replacementMap = new Map();
    repData.forEach((row: any) => {
      if (row.key !== undefined && row.value !== undefined) {
        replacementMap.set(String(row.key), String(row.value));
      }
    });

    const zip = new JSZip();
    const reportLines: string[] = [];
    
    // @ts-ignore
    for (const file of targetFiles as File[]) {
      const fileName = file.name;
      const fileBuffer = await file.arrayBuffer();
      
      const wb = XLSX.read(fileBuffer, { type: 'array' });
      let fileReplacements = 0;

      for (const sheetName of wb.SheetNames) {
        const sheet = wb.Sheets[sheetName];
        if (!sheet['!ref']) continue;
        
        const range = XLSX.utils.decode_range(sheet['!ref']);
        let sheetReplacements = 0;

        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = range.s.c; c <= range.e.c; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r, c });
            const cell = sheet[cellAddr];

            if (cell && cell.v !== undefined) {
              let newValue = String(cell.v);
              let originalValue = newValue;
              
              for (const [key, value] of replacementMap) {
                if (mode === 'full') {
                  if (newValue === key) {
                    newValue = value;
                    sheetReplacements++;
                    break;
                  }
                } else {
                  if (newValue.includes(key)) {
                     const regex = new RegExp(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
                     if (regex.test(newValue)) {
                         newValue = newValue.replace(regex, value);
                         sheetReplacements++;
                     }
                  }
                }
              }

              if (newValue !== originalValue) {
                sheet[cellAddr].v = newValue;
                if (cell.t === 'n') sheet[cellAddr].t = 's';
              }
            }
          }
        }
        fileReplacements += sheetReplacements;
        reportLines.push(`File: ${fileName} | Sheet: ${sheetName} | Replaced: ${sheetReplacements}`);
      }

      const outBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      zip.file(`replaced_${fileName}`, outBuffer);
    }

    const reportContent = reportLines.join('\n');
    zip.file('report.txt', reportContent);

    const zipContent = await zip.generateAsync({ type: 'blob' });
    const arrayBuffer = await zipContent.arrayBuffer();

    return c.body(arrayBuffer, 200, {
      'Content-Type': 'application/zip',
      'Content-Disposition': 'attachment; filename="processed_files.zip"',
      'X-Report-Header': encodeURIComponent(reportContent)
    });

  } catch (err: any) {
    console.error(err);
    return c.json({ error: err.message }, 500);
  }
});

// 2. 最后定义静态文件服务 (作为 fallback)
app.use('/*', serveStatic({ root: './public' }));

export default app;
