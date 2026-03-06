function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

self.addEventListener("message", async (event) => {
  const message = event.data || {};
  if (message.action !== "convert-blob") {
    return;
  }

  try {
    const blob = message.blob;
    if (!(blob instanceof Blob)) {
      throw new Error("worker 未收到有效 Blob");
    }

    const buffer = await blob.arrayBuffer();
    const mimeType = blob.type || "application/octet-stream";
    const base64Data = arrayBufferToBase64(buffer);

    self.postMessage({
      action: "image-ready",
      token: message.token,
      base64Data,
      mimeType,
      batchIndex: message.batchIndex ?? 0,
      imageIndex: message.imageIndex ?? 0,
      jobId: message.jobId
    });
  } catch (error) {
    self.postMessage({
      action: "image-error",
      token: message.token,
      batchIndex: message.batchIndex ?? 0,
      imageIndex: message.imageIndex ?? 0,
      jobId: message.jobId,
      error: error instanceof Error ? error.message : String(error)
    });
  }
});
