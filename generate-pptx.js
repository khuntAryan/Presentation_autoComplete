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
    console.log(`✅ Template file exists at: ${templatePath}`);

    const automizer = new Automizer({
      templateDir: path.join(__dirname, 'templates'),
      outputDir: path.join(__dirname, 'output'),
      removeExistingSlides: true, // ✅ Remove old slides from root presentation
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

    const userContent = {
      1: {
        '{{CONTENT_1_SLIDE_1}}': 'Presentation ',
      },
      2: {
        '{{CONTENT_1_SLIDE_2}}': 'Slide 2 - Primary section',
        '{{CONTENT_2_SLIDE_2}}': 'Slide 2 - Supporting data',
      },
      3: {
        '{{CONTENT_1_SLIDE_3}}': 'Slide 3 - Overview block',
        '{{CONTENT_4_SLIDE_3}}': 'Slide 3 - Extra insights',
        '{{HEADER_2_SLIDE_3}}': 'Slide 3 - Section Header 2',
        '{{HEADER_3_SLIDE_3}}': 'Slide 3 - Section Header 3',
      },
      4: {
        '{{CONTENT_1_SLIDE_4}}': 'Slide 4 - Introduction',
        '{{CONTENT_2_SLIDE_4}}': 'Slide 4 - Details block',
        '{{CONTENT_3_SLIDE_4}}': 'Slide 4 - Data analysis',
      },
      5: {
        '{{CONTENT_1_SLIDE_5}}': 'Slide 5 - Key points',
        '{{CONTENT_2_SLIDE_5}}': 'Slide 5 - Visual stats',
        '{{CONTENT_3_SLIDE_5}}': 'Slide 5 - Explanation section',
        '{{CONTENT_4_SLIDE_5}}': 'Slide 5 - Wrap-up',
        '{{CONTENT_5_SLIDE_5}}': 'Slide 5 - Final remarks',
      },
      6: {
        '{{CONTENT_1_SLIDE_6}}': 'Slide 6 - Start content',
        '{{CONTENT_2_SLIDE_6}}': 'Slide 6 - Middle section',
        '{{CONTENT_3_SLIDE_6}}': 'Slide 6 - Concluding block',
        '{{CONTENT_4_SLIDE_6}}': 'Slide 6 - Summary',
        '{{CONTENT_5_SLIDE_6}}': 'Slide 6 - Additional notes',
      },
      7: {
        '{{CONTENT_1_SLIDE_7}}': 'Slide 7 - Overview',
        '{{CONTENT_2_SLIDE_7}}': 'Slide 7 - Deep dive',
      },
      8: {
        '{{CONTENT_1_SLIDE_8}}': 'Slide 8 - High-level ideas',
        '{{CONTENT_2_SLIDE_8}}': 'Slide 8 - Trends',
        '{{CONTENT_3_SLIDE_8}}': 'Slide 8 - Use cases',
      },
      9: {
        '{{CONTENT_1_SLIDE_9}}': 'Slide 9 - Problems addressed',
        '{{CONTENT_2_SLIDE_9}}': 'Slide 9 - Solution architecture',
        '{{CONTENT_3_SLIDE_9}}': 'Slide 9 - Key takeaways',
      },
      10: {
        '{{CONTENT_1_SLIDE_10}}': 'Slide 10 - Summary points',
      },
    };
    

    console.log('✅ User content loaded for testing');

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