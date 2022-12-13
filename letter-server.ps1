# Import the required modules
Import-Module WebAdministration
Import-Module Microsoft.Office.Interop.Word
Import-Module System.Windows.Forms

# Prompt the user for the port number to use for the website
$port = [System.Windows.Forms.InputBox]::Show("Enter the port number for the website (default is 8080):", "LetterServer", "8080")
if ([string]::IsNullOrEmpty($port)) { $port = 8080 }

# Prompt the user for the physical path of the website
$physicalPath = [System.Windows.Forms.InputBox]::Show("Enter the physical path of the website (default is C:\LetterServer):", "LetterServer", "C:\LetterServer")
if ([string]::IsNullOrEmpty($physicalPath)) { $physicalPath = "C:\LetterServer" }

# Create a new website on the specified port
New-Website -Name "LetterServer" -Port $port -PhysicalPath $physicalPath

# Create a default page for the website
Set-Content "$physicalPath\index.html" (Get-Content .\index.html)

# Create a page to display the merged letter
Set-Content "$physicalPath\letter.html" (Get-Content .\letter.html)

# Prompt the user for the path to the Word template file
$templateFile = [System.Windows.Forms.InputBox]::Show("Enter the path to the Word template file:", "LetterServer", "C:\LetterTemplate.dotx")

# Open the Word template file
$word = New-Object -ComObject Word.Application
$template = $word.Documents.Open($templateFile)

# Replace the placeholder text in the template with the user's letter
$template.Content.Find.Execute("<letter>") | Out-Null
$template.Content.Text = $env:letter

# Save the merged letter as a new Word document
$fileName = "C:\Letter_$(Get-Date -Format y

