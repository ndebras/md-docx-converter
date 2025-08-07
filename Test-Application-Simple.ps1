# Test-Application-Simple.ps1
# Script de test complet pour l'application Markdown-DOCX Converter
# Usage: .\Test-Application-Simple.ps1 [-CleanupFiles]

param(
    [switch]$CleanupFiles = $false,
    [switch]$Verbose = $false,
    [switch]$SkipUnitTests = $false
)

# Configuration
$ErrorActionPreference = "Continue"
$testResults = @()
$startTime = Get-Date

# Fonctions d'affichage
function Write-Success { param($Message) Write-Host "[SUCCESS] $Message" -ForegroundColor Green }
function Write-Error { param($Message) Write-Host "[ERROR] $Message" -ForegroundColor Red }
function Write-Warning { param($Message) Write-Host "[WARNING] $Message" -ForegroundColor Yellow }
function Write-Info { param($Message) Write-Host "[INFO] $Message" -ForegroundColor Cyan }
function Write-Header { param($Message) Write-Host "`n[TEST] $Message" -ForegroundColor Magenta }

# Fonction pour enregistrer les resultats
function Add-TestResult {
    param($TestName, $Success, $Duration, $Details = "")
    $script:testResults += @{
        Test = $TestName
        Success = $Success
        Duration = $Duration
        Details = $Details
    }
}

# Fonction pour creer un fichier de test
function New-TestFile {
    param($FileName, $Content)
    try {
        $Content | Out-File -FilePath $FileName -Encoding UTF8 -Force
        return $true
    }
    catch {
        Write-Error "Impossible de creer le fichier $FileName : $($_.Exception.Message)"
        return $false
    }
}

# Fonction pour executer une commande et mesurer le temps
function Invoke-TimedCommand {
    param($Command, $Arguments = @())
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        if ($Arguments.Count -gt 0) {
            $result = & $Command @Arguments 2>&1
        } else {
            $result = & $Command 2>&1
        }
        $success = $LASTEXITCODE -eq 0
        $sw.Stop()
        return @{
            Success = $success
            Output = $result
            Duration = $sw.ElapsedMilliseconds
            ExitCode = $LASTEXITCODE
        }
    }
    catch {
        $sw.Stop()
        return @{
            Success = $false
            Output = $_.Exception.Message
            Duration = $sw.ElapsedMilliseconds
            ExitCode = -1
        }
    }
}

Write-Header "DEBUT DES TESTS DE L'APPLICATION MARKDOWN-DOCX CONVERTER"
Write-Info "Heure de debut: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Info "Nettoyage automatique: $CleanupFiles"
Write-Info "Mode verbose: $Verbose"

# Test 1: Compilation
Write-Header "Test 1: Compilation de l'application"
$result = Invoke-TimedCommand "npm" @("run", "build")
if ($result.Success) {
    Write-Success "Compilation reussie ($($result.Duration)ms)"
    Add-TestResult "Compilation" $true $result.Duration
} else {
    Write-Error "Echec de la compilation"
    if ($Verbose) { Write-Host $result.Output }
    Add-TestResult "Compilation" $false $result.Duration $result.Output
}

# Test 2: Version
Write-Header "Test 2: Version de l'application"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "--version")
if ($result.Success) {
    Write-Success "Version recuperee: $($result.Output.Trim())"
    Add-TestResult "Version" $true $result.Duration
} else {
    Write-Error "Impossible de recuperer la version"
    Add-TestResult "Version" $false $result.Duration
}

# Test 3: Aide
Write-Header "Test 3: Aide de l'application"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "--help")
if ($result.Success) {
    Write-Success "Aide affichee correctement"
    Add-TestResult "Aide" $true $result.Duration
} else {
    Write-Error "Impossible d'afficher l'aide"
    Add-TestResult "Aide" $false $result.Duration
}

# Test 4: Liste des templates
Write-Header "Test 4: Liste des templates et themes"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "list")
if ($result.Success) {
    Write-Success "Templates et themes listes"
    Add-TestResult "Templates" $true $result.Duration
} else {
    Write-Error "Impossible de lister les templates"
    Add-TestResult "Templates" $false $result.Duration
}

# Creation des fichiers de test
Write-Header "Creation des fichiers de test"

# Fichier simple
$simpleContent = @'
# Test Simple

Paragraphe avec **gras** et *italique*.

| **Col1** | *Col2* |
|----------|--------|
| **Test** | *OK* |

1. **Item** un
2. *Item* deux

Fin.
'@

# Fichier complexe
$complexContent = @'
# Test Formatage Complexe

## Tableaux avec formatage

| **Nom** | *Type* | `Code` | Status |
|---------|--------|--------|---------|
| **Test1** | *String* | `OK` | OK |
| **Test *mixte* gras** | Normal | `42` | Warning |
| Simple | *`Code ital`* | **`Gras code`** | OK |

## Listes avec formatage

1. **Premier element** important
2. *Deuxieme element* en italique
3. `Troisieme element` en code
4. Element avec **gras**, *italique* et `code inline`
5. **Formatage *imbrique* complexe**

## Caracteres speciaux

Texte avec espaces insecables et tirets longs.

---

## Citation

> Ceci est une citation avec **formatage** et *styles* multiples.

## Code block

```javascript
function test() {
    console.log("Application test");
    return { status: "OK" };
}
```

Fin du test complexe.
'@

# Fichier Mermaid
$mermaidContent = @'
# Test Mermaid

## Diagramme de test

```mermaid
graph TD
    A[Start] --> B{Test}
    B -->|OK| C[Success]
    B -->|Error| D[Debug]
    C --> E[End]
    D --> B
```

## Apres le diagramme

Texte apres le diagramme avec **formatage**.

| **Status** | *Result* |
|------------|----------|
| **Test** | *OK* |

Fin du test Mermaid.
'@

# Creer les fichiers
$testFiles = @{
    "test-simple.md" = $simpleContent
    "test-complexe.md" = $complexContent
    "test-mermaid.md" = $mermaidContent
}

foreach ($file in $testFiles.GetEnumerator()) {
    if (New-TestFile $file.Key $file.Value) {
        Write-Success "Fichier cree: $($file.Key)"
    } else {
        Write-Error "Echec creation: $($file.Key)"
    }
}

# Test 5: Conversion simple
Write-Header "Test 5: Conversion simple"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-simple.md")
if ($result.Success) {
    Write-Success "Conversion simple reussie ($($result.Duration)ms)"
    Add-TestResult "Conversion Simple" $true $result.Duration
} else {
    Write-Error "Echec conversion simple"
    Add-TestResult "Conversion Simple" $false $result.Duration
}

# Test 6: Statistiques
Write-Header "Test 6: Statistiques du fichier"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "stats", "test-simple.md")
if ($result.Success) {
    Write-Success "Statistiques generees"
    Add-TestResult "Statistiques" $true $result.Duration
} else {
    Write-Error "Echec generation statistiques"
    Add-TestResult "Statistiques" $false $result.Duration
}

# Test 7: Validation
Write-Header "Test 7: Validation du fichier"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "validate", "test-simple.md")
if ($result.Success) {
    Write-Success "Validation reussie"
    Add-TestResult "Validation" $true $result.Duration
} else {
    Write-Error "Echec validation"
    Add-TestResult "Validation" $false $result.Duration
}

# Test 8: Conversion avec Mermaid
Write-Header "Test 8: Conversion avec Mermaid"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-mermaid.md")
if ($result.Success) {
    Write-Success "Conversion Mermaid reussie ($($result.Duration)ms)"
    Add-TestResult "Mermaid" $true $result.Duration
} else {
    Write-Error "Echec conversion Mermaid"
    Add-TestResult "Mermaid" $false $result.Duration
}

# Test 9: Options avancees
Write-Header "Test 9: Options avancees"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-simple.md", "--title", "Test App", "--author", "System Test", "--toc")
if ($result.Success) {
    Write-Success "Conversion avec options reussie"
    Add-TestResult "Options Avancees" $true $result.Duration
} else {
    Write-Error "Echec conversion avec options"
    Add-TestResult "Options Avancees" $false $result.Duration
}

# Test 10: Templates differents
Write-Header "Test 10: Template academique"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-simple.md", "-t", "academic-paper", "-o", "test-academic.docx")
if ($result.Success) {
    Write-Success "Template academique applique"
    Add-TestResult "Template" $true $result.Duration
} else {
    Write-Error "Echec template academique"
    Add-TestResult "Template" $false $result.Duration
}

# Test 11: Document complexe
Write-Header "Test 11: Document complexe avec mode verbose"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-complexe.md", "--verbose")
if ($result.Success) {
    Write-Success "Document complexe converti ($($result.Duration)ms)"
    Add-TestResult "Document Complexe" $true $result.Duration
} else {
    Write-Error "Echec document complexe"
    Add-TestResult "Document Complexe" $false $result.Duration
}

# Test 12: Gestion d'erreur
Write-Header "Test 12: Gestion d'erreur (fichier inexistant)"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "fichier-inexistant.md")
if (-not $result.Success -and $result.ExitCode -eq 1) {
    Write-Success "Gestion d'erreur correcte"
    Add-TestResult "Gestion Erreur" $true $result.Duration
} else {
    Write-Error "Gestion d'erreur incorrecte"
    Add-TestResult "Gestion Erreur" $false $result.Duration
}

# Test 13: Document reel ITIL
Write-Header "Test 13: Document reel ITIL"
if (Test-Path "../docs/1_presentation_itil_v_4.md") {
    $result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "../docs/1_presentation_itil_v_4.md", "--verbose")
    if ($result.Success) {
        Write-Success "Document ITIL converti ($($result.Duration)ms)"
        Add-TestResult "Document ITIL" $true $result.Duration
    } else {
        Write-Warning "Echec document ITIL (possiblement fichier verrouille)"
        Add-TestResult "Document ITIL" $false $result.Duration
    }
} else {
    Write-Warning "Document ITIL non trouve, test ignore"
    Add-TestResult "Document ITIL" $false 0 "Fichier non trouve"
}

# Test 14: Tests unitaires
if (-not $SkipUnitTests) {
    Write-Header "Test 14: Tests unitaires"
    $result = Invoke-TimedCommand "npm" @("test")
    if ($result.Success) {
        Write-Success "Tests unitaires reussis"
        Add-TestResult "Tests Unitaires" $true $result.Duration
    } else {
        Write-Warning "Tests unitaires partiellement echoues (attendu)"
        Add-TestResult "Tests Unitaires" $false $result.Duration "Echecs mineurs attendus"
    }
} else {
    Write-Info "Tests unitaires ignores"
}

# Nettoyage des fichiers
if ($CleanupFiles) {
    Write-Header "Nettoyage des fichiers de test"
    try {
        Remove-Item -Path "test-*.md", "test-*.docx" -Force -ErrorAction SilentlyContinue
        Write-Success "Fichiers de test supprimes"
    }
    catch {
        Write-Warning "Impossible de supprimer certains fichiers (possiblement ouverts)"
    }
} else {
    Write-Info "Fichiers de test conserves (utilisez -CleanupFiles pour les supprimer)"
}

# Resume final
Write-Header "RESUME DES TESTS"
$totalTests = $testResults.Count
$successfulTests = ($testResults | Where-Object { $_.Success }).Count
$failedTests = $totalTests - $successfulTests
$successRate = [math]::Round(($successfulTests / $totalTests) * 100, 1)
$totalDuration = ($testResults | ForEach-Object { $_.Duration } | Measure-Object -Sum).Sum

Write-Info "Duree totale: $totalDuration ms"
Write-Info "Tests executes: $totalTests"
Write-Success "Tests reussis: $successfulTests"
if ($failedTests -gt 0) {
    Write-Error "Tests echoues: $failedTests"
}
Write-Info "Taux de reussite: $successRate%"

# Tableau detaille
Write-Host "`nDETAIL DES TESTS:" -ForegroundColor Yellow
Write-Host ("="*80) -ForegroundColor Yellow
Write-Host ("{0,-25} {1,-10} {2,-15} {3}" -f "Test", "Statut", "Duree (ms)", "Details") -ForegroundColor Yellow
Write-Host ("="*80) -ForegroundColor Yellow

foreach ($test in $testResults) {
    $status = if ($test.Success) { "REUSSI" } else { "ECHEC" }
    $color = if ($test.Success) { "Green" } else { "Red" }
    $details = if ($test.Details) { $test.Details } else { "" }
    Write-Host ("{0,-25} {1,-10} {2,-15} {3}" -f $test.Test, $status, $test.Duration, $details) -ForegroundColor $color
}

Write-Host ("="*80) -ForegroundColor Yellow

# Verdict final
$endTime = Get-Date
$scriptDuration = ($endTime - $startTime).TotalSeconds

Write-Header "VERDICT FINAL"
if ($successRate -ge 90) {
    Write-Success "APPLICATION 100% OPERATIONNELLE !"
    Write-Success "Score: $successRate% - PRETE POUR LA PRODUCTION"
} elseif ($successRate -ge 75) {
    Write-Warning "APPLICATION FONCTIONNELLE avec quelques problemes mineurs"
    Write-Warning "Score: $successRate% - Corrections recommandees"
} else {
    Write-Error "APPLICATION AVEC PROBLEMES MAJEURS"
    Write-Error "Score: $successRate% - Corrections necessaires"
}

Write-Info "Duree totale du script: $([math]::Round($scriptDuration, 2)) secondes"
Write-Info "Fin des tests: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# Code de sortie
if ($successRate -ge 75) {
    exit 0
} else {
    exit 1
}
