# Test-Application.ps1
# Script de test complet pour l'application Markdown-DOCX Converter
# Usage: .\Test-Application.ps1 [-CleanupFiles] [-Verbose] [-SkipUnitTests]

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
        Write-Error "Impossible de creer le fichier $FileName"
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

# Test 1: Compilation
Write-Header "Test 1: Compilation"
$result = Invoke-TimedCommand "npm" @("run", "build")
if ($result.Success) {
    Write-Success "Compilation reussie ($($result.Duration)ms)"
    Add-TestResult "Compilation" $true $result.Duration
} else {
    Write-Error "Echec de la compilation"
    Add-TestResult "Compilation" $false $result.Duration
}

# Test 2: Version
Write-Header "Test 2: Version"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "--version")
if ($result.Success) {
    Write-Success "Version: $($result.Output.Trim())"
    Add-TestResult "Version" $true $result.Duration
} else {
    Write-Error "Impossible de recuperer la version"
    Add-TestResult "Version" $false $result.Duration
}

# Test 3: Aide
Write-Header "Test 3: Aide"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "--help")
if ($result.Success) {
    Write-Success "Aide OK"
    Add-TestResult "Aide" $true $result.Duration
} else {
    Write-Error "Impossible d'afficher l'aide"
    Add-TestResult "Aide" $false $result.Duration
}

# Test 4: Templates
Write-Header "Test 4: Templates"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "list")
if ($result.Success) {
    Write-Success "Templates listes"
    Add-TestResult "Templates" $true $result.Duration
} else {
    Write-Error "Impossible de lister les templates"
    Add-TestResult "Templates" $false $result.Duration
}

# Creation des fichiers de test
Write-Header "Creation des fichiers de test"

$simpleContent = @'
# Test Simple

Paragraphe avec **gras** et *italique*.

| **Col1** | *Col2* |
|----------|--------|
| **Test** | *OK* |

1. **Item** un
2. *Item* deux
'@

$mermaidContent = @'
# Test Mermaid

```mermaid
graph TD
    A[Start] --> B[Test]
    B --> C[End]
```

Texte apres le diagramme.
'@

# Creer les fichiers
New-TestFile "test-simple.md" $simpleContent | Out-Null
New-TestFile "test-mermaid.md" $mermaidContent | Out-Null

# Test 5: Conversion simple
Write-Header "Test 5: Conversion simple"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-simple.md")
if ($result.Success) {
    Write-Success "Conversion simple OK ($($result.Duration)ms)"
    Add-TestResult "Conversion Simple" $true $result.Duration
} else {
    Write-Error "Echec conversion simple"
    Add-TestResult "Conversion Simple" $false $result.Duration
}

# Test 6: Conversion Mermaid
Write-Header "Test 6: Conversion Mermaid"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-mermaid.md")
if ($result.Success) {
    Write-Success "Conversion Mermaid OK ($($result.Duration)ms)"
    Add-TestResult "Mermaid" $true $result.Duration
} else {
    Write-Error "Echec conversion Mermaid"
    Add-TestResult "Mermaid" $false $result.Duration
}

# Test 7: Options avancees
Write-Header "Test 7: Options avancees"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "test-simple.md", "--title", "Test", "--author", "System")
if ($result.Success) {
    Write-Success "Options avancees OK"
    Add-TestResult "Options" $true $result.Duration
} else {
    Write-Error "Echec options avancees"
    Add-TestResult "Options" $false $result.Duration
}

# Test 8: Gestion erreur
Write-Header "Test 8: Gestion erreur"
$result = Invoke-TimedCommand "node" @("dist/cli/index.js", "convert", "inexistant.md")
if (-not $result.Success) {
    Write-Success "Gestion erreur OK"
    Add-TestResult "Erreurs" $true $result.Duration
} else {
    Write-Error "Gestion erreur incorrecte"
    Add-TestResult "Erreurs" $false $result.Duration
}

# Test 9: Tests unitaires
if (-not $SkipUnitTests) {
    Write-Header "Test 9: Tests unitaires"
    $result = Invoke-TimedCommand "npm" @("test")
    if ($result.Success) {
        Write-Success "Tests unitaires OK"
        Add-TestResult "Tests Unitaires" $true $result.Duration
    } else {
        Write-Warning "Tests unitaires partiels"
        Add-TestResult "Tests Unitaires" $false $result.Duration
    }
}

# Nettoyage
if ($CleanupFiles) {
    Write-Header "Nettoyage"
    Remove-Item -Path "test-*.md", "test-*.docx" -Force -ErrorAction SilentlyContinue
    Write-Success "Fichiers supprimes"
}

# Resume
Write-Header "RESUME"
$totalTests = $testResults.Count
$successfulTests = ($testResults | Where-Object { $_.Success }).Count
$successRate = [math]::Round(($successfulTests / $totalTests) * 100, 1)

Write-Info "Tests executes: $totalTests"
Write-Success "Tests reussis: $successfulTests"
Write-Info "Taux de reussite: $successRate%"

# Verdict
if ($successRate -ge 90) {
    Write-Success "APPLICATION OPERATIONNELLE !"
} elseif ($successRate -ge 75) {
    Write-Warning "Application fonctionnelle avec quelques problemes"
} else {
    Write-Error "Problemes majeurs detectes"
}

$endTime = Get-Date
$duration = ($endTime - $startTime).TotalSeconds
Write-Info "Duree totale: $([math]::Round($duration, 1)) secondes"

# Code de sortie
if ($successRate -ge 75) {
    exit 0
} else {
    exit 1
}
