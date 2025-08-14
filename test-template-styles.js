const { DocumentTemplates } = require('./dist/styles/document-templates.js');

console.log('=== TEST DES STYLES DU TEMPLATE MODERNE ===\n');

try {
  const template = DocumentTemplates.createModernTemplate();

  console.log('1. Structure complète du template:');
  console.log(JSON.stringify(template, null, 2));
  console.log();

  console.log('2. Style par défaut du document:');
  if (template.styles?.default?.document) {
    console.log('   Run:', template.styles.default.document.run);
    console.log('   Paragraph:', template.styles.default.document.paragraph);
  }
  console.log();

  console.log('3. Styles de titres:');
  if (template.styles?.headings) {
    Object.keys(template.styles.headings).forEach(headingLevel => {
      const headingStyle = template.styles.headings[headingLevel];
      console.log(`   ${headingLevel}:`, {
        run: headingStyle.run,
        paragraph: headingStyle.paragraph
      });
    });
  }
  console.log();

} catch (error) {
  console.log('Erreur lors du test:', error.message);
  console.log('Stack:', error.stack);
}

console.log('=== FIN DU TEST ===');
