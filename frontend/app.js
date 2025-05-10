document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(e.target);
  const res = await fetch('/upload-pptx', { method: 'POST', body: formData });
  const message = await res.text();
  document.getElementById('uploadStatus').textContent = message;
});

document.getElementById('processBtn').addEventListener('click', async () => {
  document.getElementById('processStatus').textContent = "Processing, please wait...";
  document.getElementById('promptContainer').classList.add('hidden');
  
  const res = await fetch('/process-pptx', { method: 'POST' });
  const message = await res.text();
  document.getElementById('processStatus').textContent = message;
  
  if (res.ok) {
    // Show AI prompt after successful processing
    try {
      const promptRes = await fetch('/get-ai-prompt');
      if (promptRes.ok) {
        const promptText = await promptRes.text();
        document.getElementById('promptContent').textContent = promptText;
        document.getElementById('promptContainer').classList.remove('hidden');
      }
    } catch (err) {
      console.error("Error fetching prompt:", err);
    }
  }
});

// Copy prompt to clipboard
document.getElementById('copyPromptBtn').addEventListener('click', async () => {
  const promptText = document.getElementById('promptContent').textContent;
  try {
    await navigator.clipboard.writeText(promptText);
    const feedback = document.getElementById('copyFeedback');
    feedback.style.display = 'inline';
    setTimeout(() => {
      feedback.style.display = 'none';
    }, 2000);
  } catch (err) {
    console.error("Error copying to clipboard:", err);
  }
});

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

document.getElementById('generateBtn').addEventListener('click', async () => {
  document.getElementById('generateStatus').textContent = "Generating, please wait...";
  document.getElementById('previewContainer').classList.add('hidden');
  const res = await fetch('/generate-pptx', { method: 'POST' });
  const message = await res.text();
  document.getElementById('generateStatus').textContent = message;

  // Check if file exists before showing preview
  const fileCheckResponse = await fetch('/check-file');
  if (fileCheckResponse.ok && (await fileCheckResponse.json()).exists) {
    document.getElementById('previewContainer').classList.remove('hidden');
  }
});
