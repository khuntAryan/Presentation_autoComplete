import { Automizer, ModifyTextHelper } from 'pptx-automizer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

// Import the mapping function
import { mapContent } from './services/mapContent.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function fillPresentation() {
  try {
    const templateFilename = 'preprocessed_sample.pptx';
    const templatePath = path.join(__dirname, 'templates', templateFilename);

    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template file not found: ${templatePath}`);
    }
    console.log(`✅ Template file exists at: ${templatePath}`);

    const automizer = new Automizer({
      templateDir: path.join(__dirname, 'templates'),
      outputDir: path.join(__dirname, 'output'),
      removeExistingSlides: true,
    });

    const pres = automizer
      .loadRoot(templateFilename)
      .load(templateFilename, 'myTemplate');

    console.log('✅ Template loaded successfully');

    const creationIds = await pres.setCreationIds();
    const myTemplate = creationIds.find(t => t.name === 'myTemplate' || t.name === '');
    if (!myTemplate) throw new Error('Template "myTemplate" not found in creationIds');

    const totalSlides = myTemplate.slides.length;
    console.log(`✅ Template has ${totalSlides} slides`);

    // Dynamically load mapped content using your mapping service
    const userContent = mapContent(
      path.join(__dirname, 'data', 'mapped-content.json'),
      path.join(__dirname, 'data', 'user-content.json')
    );

    console.log('✅ User content loaded from mapping service');

    for (let slideNum = 1; slideNum <= totalSlides; slideNum++) {
      const slideContent = userContent[slideNum] || {};

      pres.addSlide('myTemplate', slideNum, async (slide) => {
        const elements = await slide.getAllTextElementIds();
        console.log(`\n🔍 Slide ${slideNum} has ${elements.length} elements`);

        for (const elementId of elements) {
          slide.modifyElement(elementId, [
            async (element) => {
              try {
                const textContent = element.textContent || '';
                let replaced = false;

                for (const [placeholder, value] of Object.entries(slideContent)) {
                  if (textContent.includes(placeholder)) {
                    console.log(`✅ Replacing ${placeholder} in element ${elementId}`);
                    element = ModifyTextHelper.setText(value)(element);
                    replaced = true;
                    break;
                  }
                }

                if (!replaced) {
                  console.log(`ℹ️ No matching content for element ${elementId}, leaving untouched`);
                }

                return element;
              } catch (err) {
                console.error(`❌ Error processing element ${elementId}:`, err);
                return element;
              }
            }
          ]);
        }
      });
    }

    const outputFile = `filled-${Date.now()}.pptx`;
    await pres.write(outputFile);
    console.log(`\n✅ Successfully generated: ${outputFile}`);

  } catch (error) {
    console.error('❌ Error generating presentation:', error);
    throw error;
  }
}

fillPresentation()
  .then(() => console.log('\n🎉 Presentation generation completed'))
  .catch(error => console.error(`❌ Failed to generate presentation: ${error.message}`));
