
import { Hono } from 'hono';
import { serveStatic } from 'hono/cloudflare-pages';
// 显式引入 ESM 版本，避开 CommonJS 的动态 require 问题
import * as XLSX from 'xlsx/xlsx.mjs';
import JSZip from 'jszip';

// 定义环境变量类型
type Bindings = {
  AUTH_PASSWORD: string
}

const app = new Hono<{ Bindings: Bindings }>();

// --- 1. 验证密码的 API ---
app.post('/auth', async (c) => {
  try {
    const body = await c.req.parseBody();
    const password = body['password'];
    // 获取环境变量中的密码，默认为 admin
    const correctPassword = c.env.AUTH_PASSWORD || "admin";

    if (password === correctPassword) {
      return c.json({ success: true });
    } else {
      return c.json({ success: false, error: 'Incorrect password' }, 401);
    }
  } catch (e) {
    return c.json({ success: false, error: 'Auth error' }, 500);
  }
});

// --- 2. 核心处理逻辑 (带鉴权 + 完整处理) ---
app.post('/process', async (c) => {
  // A. 鉴权检查
  const authHeader = c.req.header('Authorization');
  const correctPassword = c.env.AUTH_PASSWORD || "admin";
  
  if (!authHeader || authHeader !== `Bearer ${correctPassword}`) {
    return c.json({ error: 'Unauthorized: Invalid Password' }, 401);
  }

  try {
    // B. 解析上传文件
    const body = await c.req.parseBody();
    
    // 获取文件
    let targetFiles = body['targets'];
    const replacementFile = body['replacement'];
    const mode = body['mode'] as string;

    if (!targetFiles || !replacementFile) {
      return c.json({ error: 'Missing files' }, 400);
    }

    if (!Array.isArray(targetFiles)) {
      targetFiles = [targetFiles];
    }

    // C. 读取对照表 (Replacement Map)
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

    // D. 处理目标文件并打包 ZIP
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

app.use('/*', serveStatic({ root: './public' }));

export default app;
