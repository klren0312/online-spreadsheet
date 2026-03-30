import Database from 'better-sqlite3'
import path from 'path'
import fs from 'fs'
import { fileURLToPath } from 'url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

let db: Database.Database

export function initDatabase() {
  const dbPath = path.join(__dirname, '..', 'data', 'spreadsheet.db')

  const dataDir = path.dirname(dbPath)
  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true })
  }

  db = new Database(dbPath)

  db.exec(`
    CREATE TABLE IF NOT EXISTS documents (
      name TEXT PRIMARY KEY,
      data BLOB,
      updated_at INTEGER DEFAULT (strftime('%s', 'now'))
    )
  `)

  console.log('Database initialized')
}

export function saveDocument(name: string, data: Uint8Array) {
  if (!db) return

  const stmt = db.prepare(`
    INSERT INTO documents (name, data, updated_at)
    VALUES (?, ?, strftime('%s', 'now'))
    ON CONFLICT(name) DO UPDATE SET data = ?, updated_at = strftime('%s', 'now')
  `)

  stmt.run(name, Buffer.from(data), Buffer.from(data))
}

export function loadDocument(name: string): Uint8Array | null {
  if (!db) return null

  const stmt = db.prepare('SELECT data FROM documents WHERE name = ?')
  const row = stmt.get(name) as { data: Buffer } | undefined

  if (row && row.data) {
    return new Uint8Array(row.data)
  }

  return null
}
