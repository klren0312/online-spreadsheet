import { Server as HttpServer } from 'http'
import { WebSocketServer, WebSocket } from 'ws'
import * as Y from 'yjs'
import * as syncProtocol from 'y-protocols/sync.js'
import * as awarenessProtocol from 'y-protocols/awareness.js'
import * as encoding from 'lib0/encoding.js'
import * as decoding from 'lib0/decoding.js'
import { saveDocument, loadDocument } from './database.js'

const messageSync = 0
const messageAwareness = 1

const docs = new Map<string, Y.Doc>()
const awarenessMap = new Map<string, awarenessProtocol.Awareness>()
const documentConnections = new Map<string, Set<WebSocket>>()
const clientDocMap = new Map<WebSocket, string>()

function getYDoc(docname: string): Y.Doc {
  let doc = docs.get(docname)
  if (!doc) {
    doc = new Y.Doc()
    const savedData = loadDocument(docname)
    if (savedData) {
      Y.applyUpdate(doc, savedData)
    }
    docs.set(docname, doc)

    const currentDoc = doc
    currentDoc.on('update', (_update: Uint8Array) => {
      saveDocument(docname, Y.encodeStateAsUpdate(currentDoc))
    })
  }
  return doc
}

function getAwareness(docname: string): awarenessProtocol.Awareness {
  let awareness = awarenessMap.get(docname)
  if (!awareness) {
    const doc = getYDoc(docname)
    awareness = new awarenessProtocol.Awareness(doc)

    awareness.on('update', ({ added, updated, removed }: any) => {
      const changedClients = added.concat(updated, removed)
      const encoder = encoding.createEncoder()
      encoding.writeVarUint(encoder, messageAwareness)
      encoding.writeVarUint8Array(
        encoder,
        awarenessProtocol.encodeAwarenessUpdate(awareness!, changedClients)
      )
      const message = encoding.toUint8Array(encoder)

      broadcast(docname, message, null)
    })

    awarenessMap.set(docname, awareness)
  }
  return awareness
}

function broadcast(docname: string, message: Uint8Array, exclude: WebSocket | null) {
  const connections = documentConnections.get(docname)
  if (connections) {
    connections.forEach(client => {
      if (client !== exclude && client.readyState === WebSocket.OPEN) {
        client.send(message)
      }
    })
  }
}

function handleMessage(ws: WebSocket, docname: string, message: Uint8Array) {
  const doc = getYDoc(docname)
  const awareness = getAwareness(docname)
  const decoder = new decoding.Decoder(message)
  const messageType = decoding.readVarUint(decoder)

  switch (messageType) {
    case messageSync: {
      const encoder = encoding.createEncoder()
      encoding.writeVarUint(encoder, messageSync)
      syncProtocol.readSyncMessage(decoder, encoder, doc, ws)

      if (encoding.length(encoder) > 1) {
        ws.send(encoding.toUint8Array(encoder))
      }

      broadcast(docname, message, ws)
      break
    }
    case messageAwareness: {
      awarenessProtocol.applyAwarenessUpdate(
        awareness,
        decoding.readVarUint8Array(decoder),
        ws
      )
      break
    }
  }
}

export function setupWebSocket(server: HttpServer) {
  const wss = new WebSocketServer({ server })

  wss.on('connection', (ws, req) => {
    const url = new URL(req.url || '', 'http://localhost')
    const docname = url.searchParams.get('doc') || 'default'

    const doc = getYDoc(docname)
    const awareness = getAwareness(docname)

    if (!documentConnections.has(docname)) {
      documentConnections.set(docname, new Set())
    }
    documentConnections.get(docname)!.add(ws)
    clientDocMap.set(ws, docname)

    const encoder = encoding.createEncoder()
    encoding.writeVarUint(encoder, messageSync)
    syncProtocol.writeSyncStep1(encoder, doc)
    ws.send(encoding.toUint8Array(encoder))

    ws.on('message', (data: Buffer) => {
      handleMessage(ws, docname, new Uint8Array(data))
    })

    ws.on('close', () => {
      const connections = documentConnections.get(docname)
      if (connections) {
        connections.delete(ws)
      }
      clientDocMap.delete(ws)
    })
  })

  console.log('WebSocket server ready')
}
