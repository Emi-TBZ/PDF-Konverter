# Pop-Up Fenster für den Benutzerpfad
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$default_documents_path = 'C:\doc2pdf\Word'
$default_output_path = 'C:\doc2pdf\PDF'

$documents_path = [Microsoft.VisualBasic.Interaction]::InputBox(
    'Geben Sie den Pfad zu den Word-Dateien ein (Standard ist C:\doc2pdf\Word):',
    'Eingabe erforderlich',
    $default_documents_path
)

if (-not (Test-Path -Path $documents_path)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Der eingegebene Pfad $documents_path existiert nicht.",
        "Fehler",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

$output_path = [Microsoft.VisualBasic.Interaction]::InputBox(
    'Geben Sie den Pfad ein, wo die PDFs gespeichert werden sollen (Standard ist C:\doc2pdf\PDF):',
    'Eingabe erforderlich',
    $default_output_path
)

if (-not (Test-Path -Path $output_path)) {
    New-Item -ItemType Directory -Path $output_path
    [System.Windows.Forms.MessageBox]::Show(
        "Das Zielverzeichnis $output_path wurde erstellt.",
        "Information",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

# Bestätigung vor der Konvertierung
$result = [System.Windows.Forms.MessageBox]::Show(
    "Möchten Sie die Word-Dateien aus $documents_path wirklich als PDFs in $output_path speichern?",
    "Bestätigung",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
    Write-Host "Vorgang abgebrochen." -ForegroundColor Yellow
    exit
}

# Word-Anwendung starten
$word_app = New-Object -ComObject Word.Application

# Dieser Filter sucht nach .doc und .docx Dateien
$size_report = @()
Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {
    $old_size = ($_.Length / 1MB).ToString("F2")

    $document = $word_app.Documents.Open($_.FullName)

    # PDF-Dateiname im Zielverzeichnis erstellen
    $pdf_filename = "$output_path\$($_.BaseName).pdf"

    $document.SaveAs([ref] $pdf_filename, [ref] 17)

    $document.Close()

    $new_file_info = Get-Item -Path $pdf_filename
    $new_size = ($new_file_info.Length / 1MB).ToString("F2")

    $size_report += "Datei: $($_.Name) - Alte Größe: ${old_size} MB, Neue Größe: ${new_size} MB"
    Write-Host "Konvertiert: $($_.Name) - Alte Größe: ${old_size} MB, Neue Größe: ${new_size} MB"
}

$word_app.Quit()

# Abschluss-Pop-Up
[System.Windows.Forms.MessageBox]::Show(
    "Die Konvertierung wurde abgeschlossen. PDFs sind im Verzeichnis $output_path gespeichert.",
    "Abgeschlossen",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

# Größenbericht anzeigen
[System.Windows.Forms.MessageBox]::Show(
    ($size_report -join "`n"),
    "Größen der Dateien",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

Write-Host "Konvertierung abgeschlossen. PDFs sind im Verzeichnis $output_path gespeichert." -ForegroundColor Green
