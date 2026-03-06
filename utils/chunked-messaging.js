import { CHUNKED_MESSAGE_FLAG, MAX_CHUNK_SIZE } from "./constants.js";

const inboundChunkCache = new Map();
const monitorCache = new Map();

function splitText(text, maxSize = MAX_CHUNK_SIZE) {
  const chunks = [];
  for (let i = 0; i < text.length; i += maxSize) {
    chunks.push(text.slice(i, i + maxSize));
  }
  return chunks;
}

function cleanupInbound(requestId) {
  inboundChunkCache.delete(requestId);
}

function cleanupMonitor(requestId) {
  monitorCache.delete(requestId);
}

function ensureMonitor(requestId) {
  const current = monitorCache.get(requestId);
  if (current) return current;
  const initial = {
    chunks: [],
    expectedChunks: 0,
    resolve: null,
    reject: null,
    timer: null,
    completed: false,
    result: null,
    error: null
  };
  monitorCache.set(requestId, initial);
  return initial;
}

function waitForMonitor(requestId, timeoutMs = 120000) {
  const monitor = ensureMonitor(requestId);
  if (monitor.completed) {
    if (monitor.error) {
      cleanupMonitor(requestId);
      return Promise.reject(monitor.error);
    }
    const value = monitor.result;
    cleanupMonitor(requestId);
    return Promise.resolve(value);
  }
  return new Promise((resolve, reject) => {
    monitor.resolve = resolve;
    monitor.reject = reject;
    monitor.timer = setTimeout(() => {
      cleanupMonitor(requestId);
      reject(new Error(`Chunked monitor timeout: ${requestId}`));
    }, timeoutMs);
  });
}

function resolveMonitorChunk(message) {
  const monitor = ensureMonitor(message.requestId);
  if (typeof message.totalChunks === "number" && message.totalChunks > 0) {
    monitor.expectedChunks = message.totalChunks;
  }
  if (message.type === "chunked-message-result-chunk") {
    monitor.chunks[message.chunkIndex] = message.chunk;
    return;
  }

  if (message.type === "chunked-message-result-done") {
    const text = monitor.chunks.join("");
    try {
      const parsed = JSON.parse(text);
      monitor.completed = true;
      monitor.result = parsed;
      if (monitor.timer) clearTimeout(monitor.timer);
      if (monitor.resolve) {
        monitor.resolve(parsed);
        cleanupMonitor(message.requestId);
      }
    } catch (error) {
      monitor.completed = true;
      monitor.error = error;
      if (monitor.timer) clearTimeout(monitor.timer);
      if (monitor.reject) {
        monitor.reject(error);
        cleanupMonitor(message.requestId);
      }
    }
  }
}

async function sendMessageToSender({ sender, message, runtime, context }) {
  if (context === "background" && sender?.tab?.id) {
    return chrome.tabs.sendMessage(sender.tab.id, message);
  }
  return runtime.sendMessage(message);
}

async function pushChunkedResponse({
  result,
  sender,
  runtime,
  context,
  monitorRequestId
}) {
  const serialized = JSON.stringify(result);
  const chunks = splitText(serialized, MAX_CHUNK_SIZE);
  for (let i = 0; i < chunks.length; i += 1) {
    await sendMessageToSender({
      sender,
      runtime,
      context,
      message: {
        [CHUNKED_MESSAGE_FLAG]: true,
        type: "chunked-message-result-chunk",
        requestId: monitorRequestId,
        chunkIndex: i,
        totalChunks: chunks.length,
        chunk: chunks[i],
        timestamp: Date.now()
      }
    });
  }
  await sendMessageToSender({
    sender,
    runtime,
    context,
    message: {
      [CHUNKED_MESSAGE_FLAG]: true,
      type: "chunked-message-result-done",
      requestId: monitorRequestId,
      totalChunks: chunks.length,
      done: true,
      timestamp: Date.now()
    }
  });
}

export async function sendChunkedMessage(payload, options = {}) {
  const sendMessageFn = options.sendMessageFn || ((message) => chrome.runtime.sendMessage(message));
  const timeoutMs = options.timeoutMs || 120000;

  if (options.requestIdToMonitor) {
    return waitForMonitor(options.requestIdToMonitor, timeoutMs);
  }

  const requestId = options.requestId || crypto.randomUUID();
  const serialized = JSON.stringify(payload);
  const chunks = splitText(serialized, MAX_CHUNK_SIZE);

  for (let i = 0; i < chunks.length; i += 1) {
    const ack = await sendMessageFn({
      [CHUNKED_MESSAGE_FLAG]: true,
      type: "chunked-message-chunk",
      requestId,
      chunkIndex: i,
      totalChunks: chunks.length,
      chunk: chunks[i],
      timestamp: Date.now()
    });
    if (!ack || ack[CHUNKED_MESSAGE_FLAG] !== true || ack.type !== "chunked-message-ack") {
      throw new Error(`Chunk ACK missing for ${requestId}:${i}`);
    }
  }

  const doneAck = await sendMessageFn({
    [CHUNKED_MESSAGE_FLAG]: true,
    type: "chunked-message-done",
    requestId,
    done: true,
    timestamp: Date.now()
  });

  if (!doneAck) {
    throw new Error(`Chunk done ACK missing for ${requestId}`);
  }
  if (doneAck.type === "chunked-message-error") {
    throw new Error(doneAck.error || `Chunk handler failed for ${requestId}`);
  }
  if (doneAck.type === "chunked-message-monitor" && doneAck.monitorRequestId) {
    return waitForMonitor(doneAck.monitorRequestId, timeoutMs);
  }
  return doneAck.result;
}

export function addOnChunkedMessageListener(
  handler,
  { runtime = chrome.runtime, context = "background" } = {}
) {
  const listener = (message, sender, sendResponse) => {
    if (!message || message[CHUNKED_MESSAGE_FLAG] !== true) {
      return false;
    }

    if (
      message.type === "chunked-message-result-chunk" ||
      message.type === "chunked-message-result-done"
    ) {
      resolveMonitorChunk(message);
      sendResponse({
        [CHUNKED_MESSAGE_FLAG]: true,
        type: "chunked-message-ack",
        requestId: message.requestId,
        status: "RESULT_PENDING"
      });
      return false;
    }

    if (message.type === "chunked-message-chunk") {
      const bucket =
        inboundChunkCache.get(message.requestId) ||
        {
          chunks: [],
          totalChunks: message.totalChunks || 0
        };
      bucket.chunks[message.chunkIndex] = message.chunk;
      if (message.totalChunks) {
        bucket.totalChunks = message.totalChunks;
      }
      inboundChunkCache.set(message.requestId, bucket);
      sendResponse({
        [CHUNKED_MESSAGE_FLAG]: true,
        type: "chunked-message-ack",
        requestId: message.requestId,
        chunkIndex: message.chunkIndex,
        status: "PENDING"
      });
      return false;
    }

    if (message.type === "chunked-message-done") {
      const processAsync = async () => {
        const bucket = inboundChunkCache.get(message.requestId);
        if (!bucket) {
          sendResponse({
            [CHUNKED_MESSAGE_FLAG]: true,
            type: "chunked-message-error",
            requestId: message.requestId,
            error: "Chunk bucket missing"
          });
          return;
        }

        try {
          const jsonText = bucket.chunks.join("");
          const payload = JSON.parse(jsonText);
          cleanupInbound(message.requestId);

          const result = await handler(payload, {
            sender,
            requestId: message.requestId
          });

          if (typeof result === "undefined") {
            sendResponse({
              [CHUNKED_MESSAGE_FLAG]: true,
              type: "chunked-message-ack",
              requestId: message.requestId,
              status: "DONE"
            });
            return;
          }

          const monitorRequestId = `${message.requestId}:response`;
          await pushChunkedResponse({
            result,
            sender,
            runtime,
            context,
            monitorRequestId
          });
          sendResponse({
            [CHUNKED_MESSAGE_FLAG]: true,
            type: "chunked-message-monitor",
            requestId: message.requestId,
            monitorRequestId
          });
        } catch (error) {
          cleanupInbound(message.requestId);
          sendResponse({
            [CHUNKED_MESSAGE_FLAG]: true,
            type: "chunked-message-error",
            requestId: message.requestId,
            error: error instanceof Error ? error.message : String(error)
          });
        }
      };

      processAsync();
      return true;
    }

    return false;
  };

  runtime.onMessage.addListener(listener);
  return () => runtime.onMessage.removeListener(listener);
}
