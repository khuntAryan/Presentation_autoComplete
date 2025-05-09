import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const pythonScriptPath = path.join(__dirname, 'python-preprocessor/preprocess.py');
const inputPath = path.join(__dirname, 'python-preprocessor/templates/sample2.pptx');
const outputPath = path.join(__dirname, 'python-preprocessor/output/preprocessed_sample.pptx');
const userContentPath = path.join(__dirname, 'data/user-content.json');
const mappedContentPath = path.join(__dirname, 'data/mapped-content.json');
const generatePptxScript = path.join(__dirname, 'generate-pptx.js'); // Make sure this is correct!

const command = `python3 "${pythonScriptPath}" "${inputPath}" "${outputPath}"`;

const app = express();
const port = 3000;

app.use(express.static(path.join(__dirname, 'frontend')));
app.use(express.json());

const upload = multer({ dest: 'temp/' });

app.post('/upload-pptx', upload.single('pptx'), (req, res) => {
  const tempPath = req.file.path;
  const targetPath = path.join(__dirname, 'python-preprocessor/templates/sample2.pptx');
  fs.rename(tempPath, targetPath, err => {
    if (err) {
      console.error('❌ Error moving file:', err);
      return res.status(500).send('Error saving file.');
    }
    console.log('✅ File saved to:', targetPath);
    res.send('File uploaded successfully.');
  });
});

app.post('/save-user-content', (req, res) => {
  const { bulkContent } = req.body;
  if (!bulkContent) return res.status(400).send('No content provided.');

  try {
    const parsed = parseUserContent(bulkContent);
    fs.mkdirSync(path.dirname(userContentPath), { recursive: true });
    fs.writeFileSync(userContentPath, JSON.stringify(parsed, null, 2));
    console.log('✅ User content saved:', userContentPath);
    res.send('User content parsed and saved successfully!');
  } catch (err) {
    console.error('❌ Error parsing/saving user content:', err);
    res.status(500).send('Error parsing or saving user content.');
  }
});

app.post('/process-pptx', (req, res) => {
  exec(command, (error, stdout, stderr) => {
    if (error) {
      console.error("❌ Python error:", error);
      return res.status(500).send('Error processing PPTX file.');
    }
    console.log("✅ Python stdout:", stdout);
    console.error("⚠️ Python stderr:", stderr);
    res.send('PPTX file processed successfully!');
  });
});

// NEW: Generate Final PPTX
app.post('/generate-pptx', (req, res) => {
  // 1. Run the mapping logic (mapContent.js)
  // 2. Run the generate-pptx.js script (which uses the mapping)
  exec(`node services/mapContent.js && node generate-pptx.js`, (error, stdout, stderr) => {
    if (error) {
      console.error("❌ Generation error:", error);
      return res.status(500).send('Error generating final PPTX.');
    }
    console.log("✅ Generation stdout:", stdout);
    console.error("⚠️ Generation stderr:", stderr);
    res.send('Final PPTX generated successfully! Check the output folder.');
  });
});

app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
});

// --- Helper function to parse user-pasted content ---
function parseUserContent(text) {
  const slides = text.split(/(?:^|\n)Slide\s+\d+:/gi).map(s => s.trim()).filter(Boolean);
  const result = {};
  slides.forEach((slideText, idx) => {
    const lines = slideText.split('\n').map(l => l.trim()).filter(Boolean);
    const slideKey = `slide_${idx + 1}`;
    result[slideKey] = {};

    let titleSet = false, subtitleSet = false, bullets = [], paragraph = [];
    lines.forEach(line => {
      if (!titleSet) {
        result[slideKey].title = line;
        titleSet = true;
      } else if (!subtitleSet && line && !/^[-•]/.test(line)) {
        result[slideKey].subtitle = line;
        subtitleSet = true;
      } else if (/^[-•]/.test(line)) {
        bullets.push(line.replace(/^[-•]\s*/, ''));
      } else {
        paragraph.push(line);
      }
    });
    if (bullets.length) result[slideKey].bullets = bullets;
    if (paragraph.length) result[slideKey].paragraph = paragraph.join(' ');
  });
  return result;
}
