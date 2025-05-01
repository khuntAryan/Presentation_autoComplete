import express from 'express';
import multer from 'multer';
import { Automizer, ModifyTextHelper } from 'pptx-automizer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import axios from 'axios';
import FormData from 'form-data';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ dest: 'uploads/' });

// Ensure directories exist
['templates', 'output', 'uploads'].forEach(dir => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir);
    }
});

// New function to preprocess templates with Python service
async function preprocessTemplate(templatePath) {
    try {
        const form = new FormData();
        form.append('file', fs.createReadStream(templatePath));
        
        // Call Python preprocessing service
        const response = await axios.post('http://localhost:5000/api/preprocess', form, {
            headers: {
                ...form.getHeaders(),
            },
            responseType: 'arraybuffer',
        });
        
        // Save preprocessed file
        const processedPath = path.join(__dirname, 'templates', `preprocessed_${path.basename(templatePath)}`);
        fs.writeFileSync(processedPath, response.data);
        
        return processedPath;
    } catch (error) {
        console.error('Error preprocessing template:', error);
        throw new Error(`Failed to preprocess template: ${error.message}`);
    }
}

app.post('/api/generate', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const contentData = JSON.parse(req.body.content || '{}');
        const uploadedFile = req.file;
        
        // New step: Preprocess the template with Python service
        const processedTemplatePath = await preprocessTemplate(uploadedFile.path);
        console.log(`Template preprocessed: ${processedTemplatePath}`);
        
        // Setup automizer with the preprocessed template
        const automizer = new Automizer({
            templateDir: path.dirname(processedTemplatePath),
            outputDir: path.join(__dirname, 'output'),
            removeExistingSlides: true
        });
        
        // Load the preprocessed template
        const pres = automizer.loadRoot(path.basename(processedTemplatePath));
        
        // Set creation IDs for slide information
        const creationIds = await pres.setCreationIds();
        
        if (!creationIds || creationIds.length === 0) {
            throw new Error('No templates found in the presentation');
        }
        
        // Get total slides
        const totalSlides = creationIds[0].slides.length;
        console.log(`Template has ${totalSlides} slides`);
        
        // Process each slide
        for (let slideNum = 1; slideNum <= totalSlides; slideNum++) {
            pres.addSlide(slideNum, async (slide) => {
                // Get all elements on this slide
                const elements = await slide.getAllTextElementIds();
                console.log(`Slide ${slideNum} elements:`, elements);
                
                // Get content for this slide
                const slideContent = contentData[slideNum] || {};
                
                // Apply content to each text element
                for (const elementId of elements) {
                    slide.modifyElement(elementId, [
                        (element) => {
                            const textContent = element.textContent || '';
                            
                            // Find placeholder pattern {{KEY}} and replace with content
                            for (const [key, value] of Object.entries(slideContent)) {
                                const placeholder = `{{${key}}}`;
                                if (textContent.includes(placeholder)) {
                                    console.log(`Found placeholder ${placeholder} in element ${elementId}`);
                                    return ModifyTextHelper.setText(value)(element);
                                }
                            }
                            
                            return element;
                        }
                    ]);
                }
            });
        }
        
        // Generate output file
        const outputFile = `enhanced-${Date.now()}.pptx`;
        const outputPath = path.join(__dirname, 'output', outputFile);
        await pres.write(outputFile);
        
        // Clean up temporary files
        fs.unlinkSync(uploadedFile.path);
        
        res.json({
            message: 'PowerPoint enhanced successfully',
            filePath: `/output/${outputFile}`,
            downloadUrl: `/download/${outputFile}`
        });
        
    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Download endpoint
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'output', req.params.filename);
    if (fs.existsSync(filePath)) {
        return res.download(filePath);
    }
    res.status(404).json({ error: 'File not found' });
});

app.listen(3000, () => {
    console.log('Node.js service running on port 3000');
});
