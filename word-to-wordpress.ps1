# Usage:
#
#   word-to-wordpress.ps1 -InputDir [...] -OutputDir [...] -LaTeX2WPDir [...]
#
# Description:
#
#   Goes through all Word documents in input folder (specified via -InputDir),
#   including in subfolders, and converts them to LaTeX and WordPress-ready
#   HTML. Places output in location specified via -OutputDir, mimicing the
#   subfolder structure of the input folder. For example, if the input folder
#   looks like:
#
#   input-folder/
#       a/
#           1.docx
#       b/
#           2.docx
#           c/
#               3.docx
#
#   then the output folder will look like:
#
#   output-folder/
#       a/
#           1.html
#           1.tex
#       b/
#           2.html
#           2.tex
#           c/
#               3.html
#               3.tex
#
#   The conversion is done via first using Pandoc to convert to LaTeX and then
#   using Luca Trevisan's LaTeX2WP program to convert LaTeX to HTML. The
#   location of the LaTeX2WP script must be specified via -LaTeX2WPDir. (See
#   the usage of LaTeX2WP for more info.)
#
#   Assumes dependencies (Pandoc, LaTeX2WP, Python3) are already installed.

param(
    $InputDir,
    $OutputDir,
    $LaTeX2WPDir
)

$ErrorActionPreference = "Stop";

# Create output dir, overwriting if necessary.
if (Test-Path -PathType Container -Path $OutputDir) {
    Write-Output "Output dir exists, deleting";
    Remove-Item -Force -Recurse -Path $OutputDir;
}

New-Item -ItemType Directory -Path $OutputDir;
Write-Output "Created new output dir";

# Go through all content including subfolders in input folder.
foreach ($ShortFileName in (Get-ChildItem -Path $InputDir -Recurse -Include "*.docx" -Name)) {
    Write-Output "Parsing $ShortFileName ...";

    $LongFileName = "$InputDir\$ShortFileName";
    $ShortFileNameWoutExtension = $ShortFileName.Substring(0, $ShortFileName.IndexOf(".docx"));

    # Convert from Word to LaTeX and store result in output folder.
    $TexFileName = "$OutputDir\$ShortFileNameWoutExtension.tex";
    pandoc -o $TexFileName $LongFileName;
    
    # Convert from Word to LaTeX and store result in output folder.
    $TexFileName = "$OutputDir\$ShortFileNameWoutExtension.tex";
    pandoc -s -o $TexFileName $LongFileName;

    # Convert from LaTeX to HTML and store result in output folder.
    $TmpTexFileName = "$LaTeX2WPDir\$ShortFileNameWoutExtension.tex";
    Copy-Item -Path $TexFileName -Destination $TmpTexFileName;
    python.exe "$LaTeX2WPDir\latex2wp.py" $TmpTexFileName;
    $TmpHtmlFileName = "$LaTeX2WPDir\$ShortFileNameWoutExtension.html";
    $HtmlFileName = "$OutputDir\$ShortFileNameWoutExtension.html";
    Remove-Item -Path $TmpTexFileName;
    Move-Item -Path $TmpHtmlFileName -Destination $HtmlFileName;

    # We now need to make some replacements. Replace code is based on code
    # from StackOverflow:
    # https://stackoverflow.com/questions/38466538/search-and-replace-with-powershell
    # Reasons for replacements are detailed now.
    #
    # 1. LaTeX2WP has an issue with parsing math within \( ... \) blocks,
    # and instead only recognizes math within \[ ... \]. It also seems
    # to have an issue with inline math of the form $ ... $. To fix this,
    # we replace the symbols "\(" and "\)" with "$latex" and "$" respectively.
    # Hopefully this doesn't cause any issues since it's just a simple
    # replace without any understanding of the grammar of LaTeX or HTML.
    #
    # 2. Pandoc converts "--" into "-\/-", presumably for escaping characters,
    # however for our purposes this makes "--" show up in the final WordPress
    # output as "-\/-", which is not what we want.
    $HtmlContent = Get-Content $HtmlFileName | Out-String;
    $HtmlContent | ForEach-Object{$_ -replace "\\\(", "`$latex " -replace "\\\)", "`$" -replace "-\\\/-", "--"} | Out-File $HtmlFileName;

    Write-Output "Converted $ShortFileName";
}
