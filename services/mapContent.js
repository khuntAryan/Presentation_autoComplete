import fs from 'fs';
import path from 'path';

function flattenUserContentFields(slideContent) {
  const fields = [];
  if (!slideContent) return fields;
  if (slideContent.title) fields.push(slideContent.title);
  if (slideContent.subtitle) fields.push(slideContent.subtitle);
  if (slideContent.bullets && Array.isArray(slideContent.bullets)) {
    fields.push(...slideContent.bullets);
  }
  if (slideContent.paragraph) fields.push(slideContent.paragraph);
  if (slideContent.image) fields.push(slideContent.image);
  return fields;
}

export function mapContent(
  mappedContentPath = path.join('data', 'mapped-content.json'),
  userContentPath = path.join('data', 'user-content.json')
) {
  const mappedContent = JSON.parse(fs.readFileSync(mappedContentPath, 'utf-8'));
  const userContent = JSON.parse(fs.readFileSync(userContentPath, 'utf-8'));
  const result = {};

  Object.entries(mappedContent).forEach(([slideKey, placeholders]) => {
    const slideNum = parseInt(slideKey.replace('slide_', ''), 10);
    const userFields = flattenUserContentFields(userContent[slideKey]);
    result[slideNum] = {};

    placeholders.forEach((ph, idx) => {
      result[slideNum][ph] = userFields[idx] !== undefined ? userFields[idx] : '';
    });
  });

  return result;
}


