# CLAUDE.md

## 项目概述

这是一个在线协作表格项目，支持多人实时协作编辑。

## 技术栈

- **前端**: Vue 3 + TypeScript + Vite
- **后端**: Node.js + Express + WebSocket
- **协作**: Yjs (CRDT 无冲突复制数据类型)
- **存储**: SQLite (better-sqlite3)

## 开发命令

```bash
# 安装所有依赖
npm run install:all

# 同时启动前端和后端开发服务器
npm run dev

# 单独启动后端 (端口 3001)
cd server && npm run dev

# 单独启动前端 (端口 5173)
cd client && npm run dev

# 构建前端
npm run build
```

## 项目结构

```
online-excel/
├── client/              # 前端 Vue 3 项目
│   ├── src/
│   │   ├── components/  # 表格组件
│   │   │   ├── Spreadsheet.vue      # 主表格组件
│   │   │   └── CollaboratorList.vue # 协作者列表
│   │   ├── App.vue
│   │   └── main.ts
│   └── package.json
├── server/              # 后端 Express 项目
│   ├── src/
│   │   ├── index.ts     # 入口文件
│   │   ├── websocket.ts # WebSocket 协作服务
│   │   └── database.ts  # 数据库操作
│   └── package.json
├── package.json
└── README.md
```

## 功能特性

- 实时协作：多人同时编辑，数据实时同步
- 光标显示：实时显示其他用户的光标位置
- 自动保存：数据持久化到 SQLite
- 单元格编辑、合并、复制粘贴
- 快捷键支持：方向键导航、Tab/Enter 切换
- 中文输入支持

## 代码规范

- 使用 TypeScript
- 前端使用 Vue 3 Composition API
- 保持代码简洁，避免过度抽象

## 重要提醒

- 每次功能更新后，务必同步更新 README.md 文档
