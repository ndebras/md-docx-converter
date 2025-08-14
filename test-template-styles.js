const { DocumentTemplates } = require('./dist/styles/document-templates.js');

console.log('=== TEST DES STYLES DU TEMPLATE MODERNE ===\n');

try {
  const template = DocumentTemplates.createModernTemplate();

  console.log('1. Style de cellule de tableau:');
  if (template.styles?.table?.cell) {
    console.log('   Run:', template.styles.table.cell.run);
    console.log('   Paragraph:', template.styles.table.cell.paragraph);
  } else {
    console.log('   Aucun style de cellule défini');
  }
  console.log();

  console.log('2. Style par défaut du document (pour comparaison):');
  if (template.styles?.default?.document) {
    console.log('   Run:', template.styles.default.document.run);
  }
  console.log();

  console.log('3. Vérification des tailles:');
  const tableSize = template.styles?.table?.cell?.run?.size;
  const defaultSize = template.styles?.default?.document?.run?.size;
  console.log(`   Taille police tableau: ${tableSize} (${tableSize/2}pt)`);
  console.log(`   Taille police par défaut: ${defaultSize} (${defaultSize/2}pt)`);

} catch (error) {
  console.log('Erreur lors du test:', error.message);
  console.log('Stack:', error.stack);
}

console.log('=== FIN DU TEST ===');
