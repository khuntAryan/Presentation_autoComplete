document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(e.target);
  const statusElement = document.getElementById('uploadStatus');
  
  try {
    statusElement.textContent = "Uploading and processing...";
    
    // Single request handles both upload and processing
    const uploadRes = await fetch('/upload-pptx', {
      method: 'POST',
      body: formData
    });
    
    if (!uploadRes.ok) throw new Error(await uploadRes.text());

    // Directly show AI prompt after successful processing
    statusElement.textContent = "✅ Processing complete!";
    const promptRes = await fetch('/get-ai-prompt');
    const promptText = await promptRes.text();
    
    document.getElementById('promptContent').textContent = promptText;
    document.getElementById('promptContainer').classList.remove('hidden');
    
  } catch (err) {
    statusElement.textContent = `❌ Error: ${err.message}`;
  }
});

// Attach event listener for the copy button
document.getElementById('copyPromptBtn').addEventListener('click', async () => {
  const text = document.getElementById('promptContent').textContent;
  try {
    await navigator.clipboard.writeText(text);
    // Show feedback
    const feedback = document.getElementById('copyFeedback');
    feedback.style.display = 'inline';
    setTimeout(() => {
      feedback.style.display = 'none';
    }, 1500);
  } catch (err) {
    alert("Failed to copy text. Please copy manually.");
  }
});

// Automatically save content on generate
document.getElementById('generateBtn').addEventListener('click', async () => {
  const generateStatus = document.getElementById('generateStatus');
  const bulkContent = document.getElementById('bulkContent').value;
  document.getElementById('previewContainer').classList.add('hidden');
  generateStatus.textContent = "Saving content and generating PPTX...";

  try {
    // 1. Save content automatically
    const saveRes = await fetch('/save-user-content', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bulkContent })
    });
    if (!saveRes.ok) throw new Error(await saveRes.text());

    // 2. Generate PPTX
    const genRes = await fetch('/generate-pptx', { method: 'POST' });
    const message = await genRes.text();
    generateStatus.textContent = message;

    // 3. Show preview if successful
    const fileCheck = await fetch('/check-file');
    if (fileCheck.ok && (await fileCheck.json()).exists) {
      document.getElementById('previewContainer').classList.remove('hidden');
    }
  } catch (err) {
    generateStatus.textContent = `Error: ${err.message}`;
  }
});
