# Quick-Test-Simple.ps1
# Script de test rapide pour validation de base
# Usage: .\Quick-Test-Simple.ps1 [-CleanupFiles]

param(
    [switch]$CleanupFiles = $false
)

Write-Host "[TEST] TEST RAPIDE DE L'APPLICATION" -ForegroundColor Magenta
Write-Host "=" * 50

# Test de compilation
Write-Host "`n[BUILD] Compilation..." -ForegroundColor Yellow
npm run build
if ($LASTEXITCODE -eq 0) {
    Write-Host "[SUCCESS] Compilation OK" -ForegroundColor Green
} else {
    Write-Host "[ERROR] Echec compilation" -ForegroundColor Red
    exit 1
}

# Test basique
Write-Host "`n[CONVERT] Test de conversion simple..." -ForegroundColor Yellow
@"
# Test Rapide

Paragraphe avec **gras** et *italique*.

| **Col1** | *Col2* |
|----------|--------|
| Test | OK |

1. Item un
2. Item deux
"@ | Out-File -FilePath "quick-test.md" -Encoding UTF8

node dist/cli/index.js convert quick-test.md -o quick-output.docx
if ($LASTEXITCODE -eq 0) {
    Write-Host "[SUCCESS] Conversion OK" -ForegroundColor Green
} else {
    Write-Host "[ERROR] Echec conversion" -ForegroundColor Red
    exit 1
}

# Test avec Mermaid
Write-Host "`n[MERMAID] Test Mermaid..." -ForegroundColor Yellow
@"
# Test Mermaid

```mermaid
graph TD
    A[Start] --> B[Test]
    B --> C[End]
```
"@ | Out-File -FilePath "quick-mermaid.md" -Encoding UTF8

node dist/cli/index.js convert quick-mermaid.md -o quick-mermaid.docx
if ($LASTEXITCODE -eq 0) {
    Write-Host "[SUCCESS] Mermaid OK" -ForegroundColor Green
} else {
    Write-Host "[ERROR] Echec Mermaid" -ForegroundColor Red
}

# Nettoyage
if ($CleanupFiles) {
    Remove-Item -Path "quick-*.md", "quick-*.docx" -Force -ErrorAction SilentlyContinue
    Write-Host "`n[CLEANUP] Fichiers nettoyes" -ForegroundColor Cyan
} else {
    Write-Host "`n[INFO] Fichiers conserves (quick-test.md, quick-output.docx, quick-mermaid.md, quick-mermaid.docx)" -ForegroundColor Cyan
}

Write-Host "`n[SUCCESS] TEST RAPIDE TERMINE !" -ForegroundColor Green
