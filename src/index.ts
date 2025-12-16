
import { Hono } from 'hono';
import { serveStatic } from 'hono/cloudflare-pages';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

// 定义 Cloudflare 环境变量的类型
type Bindings = {
  AUTH_PASSWORD: string
}

const app = new Hono<{ Bindings: Bindings }>();

// 1. 新增：验证密码的 API (用于前端登录时立即反馈)
app.post('/auth', async (c) => {
  const body = await c.req.parseBody();
  const password = body['password'];
  
  // 从环境变量获取密码，如果没有设置，默认是 "admin"
  const correctPassword = c.env.AUTH_PASSWORD || "admin";

  if (password === correctPassword) {
    return c.json({ success: true });
  } else {
    return c.json({ success: false, error: 'Incorrect password' }, 401);
  }
});

// 2. 修改：核心处理逻辑，增加鉴权拦截
app.post('/process', async (c) => {
  // --- 鉴权开始 ---
  const authHeader = c.req.header('Authorization');
  const correctPassword = c.env.AUTH_PASSWORD || "admin";
  
  // 前端会发送 "Bearer 你的密码"
  if (!authHeader || authHeader !== `Bearer ${correctPassword}`) {
    return c.json({ error: 'Unauthorized: Invalid Password' }, 401);
  }
  // --- 鉴权结束 ---

  try {
    const body = await c.req.parseBody();
    // ... (原本的文件处理逻辑完全保持不变) ...
    // ... 为了节省篇幅，这里省略中间重复的 Excel 处理代码 ...
    // ... 请将原本的 Excel 处理逻辑放在这里 ...
    
    // ⬇️ 仅仅为了完整性示意，这里保留原本的核心变量获取逻辑，请确保你保留了之前的完整代码
    let targetFiles = body['targets'];
    const replacementFile = body['replacement'];
    const mode = body['mode'] as string;
    
    if (!targetFiles || !replacementFile) return c.json({ error: 'Missing files' }, 400);
    if (!Array.isArray(targetFiles)) targetFiles = [targetFiles];

    // ... (此处省略几十行 Excel 处理代码) ...
    // ... 请直接复制之前运行正常的逻辑 ...

    // 假设 zip 生成完毕 (仅作示意，请使用你现有的代码)
    const zip = new JSZip(); 
    // ... 
    const zipContent = await zip.generateAsync({ type: 'blob' });
    const arrayBuffer = await zipContent.arrayBuffer();

    return c.body(arrayBuffer, 200, {
      'Content-Type': 'application/zip',
      'Content-Disposition': 'attachment; filename="processed_files.zip"',
      'X-Report-Header': encodeURIComponent("Report...")
    });

  } catch (err: any) {
    console.error(err);
    return c.json({ error: err.message }, 500);
  }
});

app.use('/*', serveStatic({ root: './public' }));

export default app;
