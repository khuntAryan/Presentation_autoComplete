import { Automizer, ModifyTextHelper } from 'pptx-automizer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function fillPresentation() {
  try {
    // Updated template filename
    const templateFilename = 'preprocessed_sample.pptx';
    const templatePath = path.join(__dirname, 'templates', templateFilename);
    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template file not found: ${templatePath}`);
    }
    console.log(`Template file exists at: ${templatePath}`);

    const automizer = new Automizer({
      templateDir: path.join(__dirname, 'templates'),
      outputDir: path.join(__dirname, 'output'),
      removeExistingSlides: true,
    });

    // Load template
    const pres = automizer
      .loadRoot(templateFilename)
      .load(templateFilename, 'myTemplate');

    console.log('Template loaded successfully');

    const creationIds = await pres.setCreationIds();
    console.log('Creation IDs:', JSON.stringify(creationIds, null, 2));

    const myTemplate = creationIds.find(t => t.name === 'myTemplate' || t.name === '');
    if (!myTemplate) {
      throw new Error('Template "myTemplate" not found in creationIds');
    }

    const totalSlides = myTemplate.slides.length;
    console.log(`Template has ${totalSlides} slides`);

    // Content mapped by slide
    const content = {
      1: {
        [`{{TITLE_SLIDE_1}}`]: "The Great Green Wall Initiative",
        [`{{SUBTITLE_SLIDE_1}}`]: "Restoring Life on Land in Africa - A Case Study in Social Engineering"
      },
      2: {
        [`{{CARD1_TITLE_SLIDE_2}}`]: "The Sahel's Challenge",
        [`{{CARD1_CONTENT_SLIDE_2}}`]: "• Once fertile region stretching 8,000km across Africa\n• Now faces severe desertification affecting 40% of land\n• Communities suffering from poverty, food insecurity, and migration\n• Climate change amplifying desertification at 3-5% annually",
        [`{{CARD2_TITLE_SLIDE_2}}`]: "The Vision of the Great Green Wall",
        [`{{CARD2_CONTENT_SLIDE_2}}`]: "• Restore 100 million hectares of degraded land by 2030\n• Create 10 million sustainable jobs for rural populations\n• Sequester 250 million tons of carbon dioxide\n• Build resilience for 250 million people in the Sahel"
      },
      3: {
        [`{{CHALLENGE_TITLE_SLIDE_3}}`]: "Community-Driven Solutions",
        [`{{CHALLENGE_BULLETS_SLIDE_3}}`]: "• Community-led restoration through agroforestry training\n• Water conservation using 'half-moon' stone barriers\n• Planting drought-resistant native species like Acacia senegal\n• Leveraging satellite imagery and AI for monitoring\n• Integrating trees with crops for better farm productivity\n• Senegal's focus on fruit trees providing both environmental restoration and income"
      }
    };

    for (let slideNum = 1; slideNum <= totalSlides; slideNum++) {
      pres.addSlide('myTemplate', slideNum, async (slide) => {
        const elements = await slide.getAllTextElementIds();
        console.log(`Slide ${slideNum} elements:`, elements);

        for (const elementId of elements) {
          slide.modifyElement(elementId, [
            async (element) => {
              try {
                const textContent = element.textContent || '';

                const slideContent = content[slideNum];
                if (slideContent) {
                  for (const [placeholder, value] of Object.entries(slideContent)) {
                    if (textContent.includes(placeholder)) {
                      console.log(`Replacing ${placeholder} in element ${elementId}`);
                      element = ModifyTextHelper.setText(value)(element);
                    }
                  }
                }

                return element;
              } catch (err) {
                console.error(`Error processing element ${elementId}:`, err);
                return element;
              }
            }
          ]);
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
