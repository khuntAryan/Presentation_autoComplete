document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(e.target);
  
  const res = await fetch('/upload-pptx', {
  method: 'POST',
  body: formData,
  });
  
  const message = await res.text();
  document.getElementById('uploadStatus').textContent = message;
  });
  
  // NEW: Handle single textarea for all content
  document.getElementById('contentForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const bulkContent = document.getElementById('bulkContent').value;
  const res = await fetch('/save-user-content', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ bulkContent })
  });
  
  const message = await res.text();
  document.getElementById('contentStatus').textContent = message;
  });
  
  document.getElementById('processBtn').addEventListener('click', async () => {
  const res = await fetch('/process-pptx', { method: 'POST' });
  const message = await res.text();
  document.getElementById('processStatus').textContent = message;
  });
  // ...existing code...

document.getElementById('generateBtn').addEventListener('click', async () => {
  const res = await fetch('/generate-pptx', { method: 'POST' });
  const message = await res.text();
  document.getElementById('generateStatus').textContent = message;
});
