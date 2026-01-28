import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import path from 'path'

// 自定义插件：处理 CSV 保存请求
const csvServer = () => ({
  name: 'csv-server',
  configureServer(server) {
    server.middlewares.use('/api/save-csv', async (req, res, next) => {
      if (req.method === 'POST') {
        let body = '';
        req.on('data', chunk => {
          body += chunk.toString();
        });
        req.on('end', () => {
          try {
            // 解析前端发来的 JSON
            const { csvContent } = JSON.parse(body);
            // 写入文件到 public/specs.csv
            const filePath = path.resolve(__dirname, 'public/specs.csv');
            fs.writeFileSync(filePath, csvContent, 'utf-8');
            
            res.statusCode = 200;
            res.end(JSON.stringify({ success: true }));
            console.log('✅ CSV 文件已更新！');
          } catch (err) {
            console.error('❌ 保存失败:', err);
            res.statusCode = 500;
            res.end(JSON.stringify({ error: err.message }));
          }
        });
      } else {
        next();
      }
    });
  }
})

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react(), csvServer()],
  server: {
    host: true, // 开启局域网访问
  }
})