import express from 'express'
import { createServer } from 'http'
import cors from 'cors'
import { setupWebSocket } from './websocket.js'
import { initDatabase } from './database.js'

const app = express()
const server = createServer(app)

app.use(cors())
app.use(express.json())

initDatabase()

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok' })
})

setupWebSocket(server)

const PORT = process.env.PORT || 3001

server.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`)
})
