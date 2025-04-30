import { Automizer, ModifyTextHelper } from 'pptx-automizer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function fillPresentation() {
  try {
    // Verify template exists
    const templatePath = path.join(__dirname, 'templates', 'test.pptx');
    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template file not found: ${templatePath}`);
    }
    console.log(`Template file exists at: ${templatePath}`);

    const automizer = new Automizer({
      templateDir: path.join(__dirname, 'templates'),
      outputDir: path.join(__dirname, 'output'),
      removeExistingSlides: true,
    });

    // Load template with explicit name
    const pres = automizer
      .loadRoot('test.pptx')
      .load('test.pptx', 'myTemplate');
    
    console.log('Template loaded successfully');
    
    // Get creation IDs with error handling
    const creationIds = await pres.setCreationIds();
    console.log('Creation IDs:', JSON.stringify(creationIds, null, 2));
    
    // Check if any templates were found
    if (!creationIds || creationIds.length === 0) {
      throw new Error('No templates found in creationIds');
    }
    
    // Find the template by name rather than assuming index 0
    const myTemplate = creationIds.find(t => t.name === 'myTemplate' || t.name === '');
    if (!myTemplate) {
      throw new Error('Template "myTemplate" not found in creationIds');
    }
    
    const totalSlides = myTemplate.slides.length;
    console.log(`Template has ${totalSlides} slides`);
    
    // Define comprehensive content for all slides
    const content = {
      1: {
        title: "The Great Green Wall Initiative",
        subtitle: "Restoring Life on Land in Africa - A Case Study in Social Engineering"
      },
      2: {
        card1_title: "The Sahel's Challenge",
        card1_content: "• Once fertile region stretching 8,000km across Africa\n• Now faces severe desertification affecting 40% of land\n• Communities suffering from poverty, food insecurity, and migration\n• Climate change amplifying desertification at 3-5% annually",
        card2_title: "The Vision of the Great Green Wall",
        card2_content: "• Restore 100 million hectares of degraded land by 2030\n• Create 10 million sustainable jobs for rural populations\n• Sequester 250 million tons of carbon dioxide\n• Build resilience for 250 million people in the Sahel"
      },
      3: {
        challenge_title: "Community-Driven Solutions",
        challenge_bullets: "• Community-led restoration through agroforestry training\n• Water conservation using 'half-moon' stone barriers\n• Planting drought-resistant native species like Acacia senegal\n• Leveraging satellite imagery and AI for monitoring\n• Integrating trees with crops for better farm productivity\n• Senegal's focus on fruit trees providing both environmental restoration and income"
      }
      
    };
    
    // Process slides
    for (let slideNum = 1; slideNum <= totalSlides; slideNum++) {
      pres.addSlide('myTemplate', slideNum, async (slide) => {
        const elements = await slide.getAllTextElementIds();
        console.log(`Slide ${slideNum} elements:`, elements);
        
        // First try to identify elements based on their current text content
        for (const elementId of elements) {
          slide.modifyElement(elementId, [
            (element) => {
              try {
                const textContent = element.textContent || '';
                
                // Slide 1 content
                if (textContent.includes('{{TITLE}}')) {
                  console.log(`Found {{TITLE}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[1].title)(element);
                }
                else if (textContent.includes('{{SUBTITLE}}')) {
                  console.log(`Found {{SUBTITLE}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[1].subtitle)(element);
                }
                
                // Slide 2 content
                else if (textContent.includes('{{CARD1_TITLE}}')) {
                  console.log(`Found {{CARD1_TITLE}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[2].card1_title)(element);
                }
                else if (textContent.includes('{{CARD1_CONTENT}}')) {
                  console.log(`Found {{CARD1_CONTENT}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[2].card1_content)(element);
                }
                else if (textContent.includes('{{CARD2_TITLE}}')) {
                  console.log(`Found {{CARD2_TITLE}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[2].card2_title)(element);
                }
                else if (textContent.includes('{{CARD2_CONTENT}}')) {
                  console.log(`Found {{CARD2_CONTENT}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[2].card2_content)(element);
                }
                
                // Slide 3 content
                else if (textContent.includes('{{CHALLENGE_TITLE}}')) {
                  console.log(`Found {{CHALLENGE_TITLE}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[3].challenge_title)(element);
                }
                else if (textContent.includes('{{CHALLENGE_BULLETS}}')) {
                  console.log(`Found {{CHALLENGE_BULLETS}} in element ${elementId}`);
                  element = ModifyTextHelper.setText(content[3].challenge_bullets)(element);
                }
                
                return element;
              } catch (err) {
                console.error(`Error processing element ${elementId}:`, err);
                return element;
              }
            }
          ]);
        }
        
        // Fallback approach using standard PowerPoint element naming
        if (slideNum === 1) {
          // Try standard title/subtitle element names
          if (elements.includes('Title 1')) {
            slide.modifyElement('Title 1', [
              ModifyTextHelper.setText(content[1].title)
            ]);
          }
          if (elements.includes('Subtitle 2')) {
            slide.modifyElement('Subtitle 2', [
              ModifyTextHelper.setText(content[1].subtitle)
            ]);
          }
        }
        else if (slideNum === 2) {
          // Try to identify card text boxes by their position in elements array
          const textBoxes = elements.filter(e => e.includes('TextBox'));
          if (textBoxes.length >= 4) {
            slide.modifyElement(textBoxes[0], [
              ModifyTextHelper.setText(content[2].card1_title)
            ]);
            slide.modifyElement(textBoxes[1], [
              ModifyTextHelper.setText(content[2].card1_content)
            ]);
            slide.modifyElement(textBoxes[2], [
              ModifyTextHelper.setText(content[2].card2_title)
            ]);
            slide.modifyElement(textBoxes[3], [
              ModifyTextHelper.setText(content[2].card2_content)
            ]);
          }
        }
        else if (slideNum === 3) {
          // Try to identify challenge bullets text boxes
          const textBoxes = elements.filter(e => e.includes('TextBox'));
          if (textBoxes.length >= 1) {
            slide.modifyElement(textBoxes[0], [
              ModifyTextHelper.setText(content[3].challenge_bullets)
            ]);
          }
        }
      });
    }

    const outputFile = `filled-${Date.now()}.pptx`;
    await pres.write(outputFile);
    console.log(`Successfully generated: ${outputFile}`);
    
  } catch (error) {
    console.error('Error generating presentation:', error);
    throw error;
  }
}

fillPresentation()
  .then(() => console.log('Presentation generation completed'))
  .catch(error => console.error(`Failed to generate presentation: ${error.message}`));
