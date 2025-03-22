#!/usr/bin/env pwsh
<#
.SYNOPSIS
    A PowerShell web server that provides an interface for writing letters and merging them with Word templates.

.DESCRIPTION
    LetterServer creates a simple web server on a specified port and provides an HTML interface
    for users to write letters. The script can merge the letter content with a specified Word
    template and save the resulting document.

.NOTES
    Author: Armin Marth
    Version: 1.1.0
    Last Updated: 2025-03-22
    Requires: PowerShell 5.1+, WebAdministration module, Microsoft.Office.Interop.Word module
#>

# Import the required modules
Import-Module WebAdministration -ErrorAction SilentlyContinue
Import-Module Microsoft.Office.Interop.Word -ErrorAction SilentlyContinue

# Check if required modules are available
$modulesAvailable = $true
if (-not (Get-Module -Name WebAdministration -ListAvailable)) {
    Write-Warning "WebAdministration module is not available. Please install it using: Install-Module -Name WebAdministration -Force"
    $modulesAvailable = $false
}

if (-not (Get-Module -Name Microsoft.Office.Interop.Word -ListAvailable)) {
    Write-Warning "Microsoft.Office.Interop.Word module is not available. Please ensure Microsoft Word is installed."
    $modulesAvailable = $false
}

if (-not $modulesAvailable) {
    Write-Error "Required modules are missing. Please install them and try again."
    exit 1
}

# Function to show input dialog
function Show-InputDialog {
    param (
        [string]$prompt,
        [string]$title,
        [string]$default
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(380, 20)
    $label.Text = $prompt
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(360, 20)
    $textBox.Text = $default
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75, 100)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(250, 100)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $default
    }
}

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Prompt the user for the port number to use for the website
$port = Show-InputDialog -prompt "Enter the port number for the website:" -title "LetterServer" -default "8080"
if ([string]::IsNullOrEmpty($port)) { $port = 8080 }

# Prompt the user for the physical path of the website
$physicalPath = Show-InputDialog -prompt "Enter the physical path of the website:" -title "LetterServer" -default "$scriptDir\WebSite"
if ([string]::IsNullOrEmpty($physicalPath)) { $physicalPath = "$scriptDir\WebSite" }

# Create the physical path directory if it doesn't exist
if (-not (Test-Path $physicalPath)) {
    New-Item -ItemType Directory -Path $physicalPath -Force | Out-Null
    Write-Host "Created directory: $physicalPath"
}

# Check if the website already exists
$existingWebsite = Get-Website -Name "LetterServer" -ErrorAction SilentlyContinue
if ($existingWebsite) {
    Write-Host "Website 'LetterServer' already exists. Removing it..."
    Remove-Website -Name "LetterServer"
}

# Create a new website on the specified port
try {
    New-Website -Name "LetterServer" -Port $port -PhysicalPath $physicalPath -ErrorAction Stop
    Write-Host "Created website 'LetterServer' on port $port"
} catch {
    Write-Error "Failed to create website: $_"
    exit 1
}

# Copy the HTML files to the website directory
try {
    Copy-Item -Path "$scriptDir\index.html" -Destination "$physicalPath\index.html" -Force
    Copy-Item -Path "$scriptDir\letter.html" -Destination "$physicalPath\letter.html" -Force
    Write-Host "Copied HTML files to $physicalPath"
} catch {
    Write-Error "Failed to copy HTML files: $_"
    exit 1
}

# Prompt the user for the path to the Word template file
$templateFile = Show-InputDialog -prompt "Enter the path to the Word template file:" -title "LetterServer" -default "$scriptDir\LetterTemplate.dotx"

# Verify the template file exists
if (-not (Test-Path $templateFile)) {
    Write-Warning "Template file not found: $templateFile"
    Write-Host "Creating a simple template file..."
    
    # Create a simple Word template
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    
    # Add some basic content to the template
    $doc.Content.Text = "Dear Sir/Madam,`n`n<letter>`n`nSincerely,`nYour Name"
    
    # Save the template
    $doc.SaveAs([ref]$templateFile, [ref]17) # 17 = wdFormatXMLTemplate
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    
    Write-Host "Created template file: $templateFile"
}

# Start the website
Start-Website -Name "LetterServer"
Write-Host "Started website 'LetterServer'"
Write-Host "Access the letter writing interface at: http://localhost:$port/"

# Function to process letter submissions
function Process-Letter {
    param (
        [string]$letterContent
    )
    
    # Open the Word template file
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    try {
        $template = $word.Documents.Open($templateFile)
        
        # Replace the placeholder text in the template with the user's letter
        if ($template.Content.Find.Execute("<letter>")) {
            $range = $template.Content.Find.Found
            $range.Text = $letterContent
        } else {
            # If placeholder not found, append the letter to the end
            $template.Content.InsertAfter("`n`n$letterContent")
        }
        
        # Generate a filename with timestamp
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $fileName = "$scriptDir\Letter_$timestamp.docx"
        
        # Save the merged letter as a new Word document
        $template.SaveAs([ref]$fileName)
        $template.Close()
        
        Write-Host "Letter saved as: $fileName"
        return $fileName
    } catch {
        Write-Error "Error processing letter: $_"
    } finally {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# Set up a listener for HTTP requests
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")

try {
    $listener.Start()
    Write-Host "HTTP listener started on port $port"
    
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        # Get the requested URL
        $url = $request.Url.LocalPath
        Write-Host "Request received for: $url"
        
        if ($request.HttpMethod -eq "POST" -and $url -eq "/letter.html") {
            # Read the form data
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $formData = $reader.ReadToEnd()
            $reader.Close()
            
            # Parse the form data to get the letter content
            $letterContent = [System.Web.HttpUtility]::UrlDecode($formData -replace "letter=", "")
            
            # Process the letter
            $fileName = Process-Letter -letterContent $letterContent
            
            # Set the letter content in the response environment
            $env:letter = $letterContent
            
            # Serve the letter.html page
            $content = Get-Content "$physicalPath\letter.html" -Raw
            $content = $content -replace '\$\(\$env:letter\)', $letterContent
            
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($content)
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
        } else {
            # Serve the requested file
            $filePath = "$physicalPath$url"
            if ($url -eq "/") { $filePath = "$physicalPath\index.html" }
            
            if (Test-Path $filePath) {
                $content = Get-Content $filePath -Raw
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($content)
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            } else {
                # Return 404 Not Found
                $response.StatusCode = 404
                $notFoundMessage = "404 - File not found: $url"
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($notFoundMessage)
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
        }
        
        $response.Close()
    }
} catch {
    Write-Error "Error in HTTP listener: $_"
} finally {
    # Stop the listener when done
    if ($listener.IsListening) {
        $listener.Stop()
        Write-Host "HTTP listener stopped"
    }
    
    # Stop the website
    Stop-Website -Name "LetterServer"
    Write-Host "Stopped website 'LetterServer'"
}
