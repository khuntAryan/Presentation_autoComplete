import { Automizer, ModifyTextHelper } from 'pptx-automizer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function fillPresentation() {
  try {
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

    // Expanded content with additional placeholders
    const content = {
      1: {
        [`{{TITLE_SLIDE_1}}`]: "The Great Green Wall Initiative",
        [`{{SUBTITLE_SLIDE_1}}`]: "Restoring Life on Land in Africa - A Case Study in Social Engineering"
      },
      2: {
        [`{{RIGHT_CONTENT_SLIDE_2_1}}`]: "The Sahel's Challenge: severe desertification threatening livelihoods.",
        [`{{LEFT_CONTENT_SLIDE_2_1}}`]: "Vision: Restore 100 million hectares, create jobs, and build resilience."
      },
      3: {
        [`{{TITLE_SLIDE_3}}`]: "Community-Driven Solutions",
        [`{{CONTENT_SLIDE_3_1}}`]: "Community-led agroforestry, water conservation with stone barriers, native species planting.",
        [`{{CONTENT_SLIDE_3_2}}`]: "Use of satellite monitoring, AI tools, and integrating trees with crops."
      },
      4: {
        [`{{TITLE_SLIDE_4}}`]: "Impact Metrics",
        [`{{CONTENT_SLIDE_4_1}}`]: "• 20 million hectares restored\n• 350,000 jobs created\n• 20 million people benefited so far",
        [`{{CONTENT_SLIDE_4_2}}`]: "Goal: Expand restoration to cover 100 million hectares by 2030."
      },
      5: {
        [`{{TITLE_SLIDE_5}}`]: "Global Partnerships",
        [`{{CONTENT_SLIDE_5_1}}`]: "Supported by African Union, UN agencies, World Bank, NGOs.",
        [`{{CONTENT_SLIDE_5_2}}`]: "Collaboration with private sector to drive financing and implementation."
      }
      // Add more slides similarly if needed
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
