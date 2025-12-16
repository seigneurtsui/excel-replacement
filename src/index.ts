import { Hono } from 'hono';
import { serveStatic } from 'hono/cloudflare-pages';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

const app = new Hono();

// 1. 提供静态资源 (前端页面)
app.use('/*', serveStatic({ root: './public' }));

// 2. 处理 Excel 替换的核心 API
app.post('/process', async (c) => {
  try {
    const body = await c.req.parseBody();
    
    // 获取上传的文件
    // Hono parseBody 对多文件上传的处理：如果是多个同名文件，会返回数组
    let targetFiles = body['targets'];
    const replacementFile = body['replacement'];
    const mode = body['mode'] as string;

    if (!targetFiles || !replacementFile) {
      return c.json({ error: 'Missing files' }, 400);
    }

    // 规范化 targets，确保它是数组
    if (!Array.isArray(targetFiles)) {
      targetFiles = [targetFiles];
    }

    // --- 读取对照表 (Replacement Map) ---
    // @ts-ignore: Hono types for file can be explicitly cast or handled
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

    // --- 准备 ZIP 文件 ---
    const zip = new JSZip();
    const reportLines: string[] = [];
    
    // --- 处理目标文件 ---
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
                    break; // 完全匹配找到后即可跳出
                  }
                } else {
                  // 部分匹配
                  if (newValue.includes(key)) {
                     // 使用正则全局替换，注意转义特殊字符
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
                // 如果原来是数字类型，修改后改为字符串类型，避免 Excel 报错
                if (cell.t === 'n') sheet[cellAddr].t = 's';
              }
            }
          }
        }
        fileReplacements += sheetReplacements;
        reportLines.push(`File: ${fileName} | Sheet: ${sheetName} | Replaced: ${sheetReplacements}`);
      }

      // 将修改后的 Excel 写入 Buffer
      const outBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      
      // 添加到 ZIP
      zip.file(`replaced_${fileName}`, outBuffer);
    }

    // --- 添加报告到 ZIP ---
    const reportContent = reportLines.join('\n');
    zip.file('report.txt', reportContent);

    // --- 生成 ZIP 并返回 ---
    const zipContent = await zip.generateAsync({ type: 'blob' });
    const arrayBuffer = await zipContent.arrayBuffer();

    return c.body(arrayBuffer, 200, {
      'Content-Type': 'application/zip',
      'Content-Disposition': 'attachment; filename="processed_files.zip"',
      'X-Report-Header': encodeURIComponent(reportContent) // 通过 Header 传递简要报告给前端展示
    });

  } catch (err: any) {
    console.error(err);
    return c.json({ error: err.message }, 500);
  }
});

export default app;
