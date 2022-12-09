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
Set-Content "$physicalPath\index.html" "<html><head><style>body{font-family:Arial,sans-serif;margin:0;padding:0;background-color:#f3f3f3}.content{max-width:500px;margin:auto;background-color:#fff;border:1px solid #ccc;border-radius:5px;padding:16px}.form{margin-top:16px}.form textarea{width:100%;box-sizing:border-box;font-size:16px;padding:12px;border:1px solid #ccc;border-radius:4px;resize:vertical}.form input[type='submit']{width:100%;background-color:#4CAF50;color:#fff;padding:12px 20px;margin:8px 0;border:none;border-radius:4px;cursor:pointer;font-size:18px}</style></head><body><div class='content'><h1>Write Your Letter</h1><form class='form' method='post' action='letter.html'><textarea name='letter' rows='10' cols='50'></textarea><br/><input type='submit' value='Submit'></form></div></body></html>"

# Create a page to display the merged letter
Set-Content "$physicalPath\letter.html" "<html><head><style>body{font-family:Arial,sans-serif;margin:0;padding:0;background-color:#f3f3f3}.content{max-width:500px;margin:auto;background-color:#fff;border:1px solid #ccc;border-radius:5px;padding:16px}pre{white-space:pre-wrap;word-wrap:break-word}</style></head><body><div class='content'><h1>Your Letter</h1><pre>$($env:letter)</pre></div></body></html>"

# Prompt the user for the path to the Word template file
$templateFile = [System.Windows.Forms.InputBox]::Show("Enter the path to the Word template file:", "LetterServer", "C:\LetterTemplate.dotx")

# Open the Word template file
$word = New-Object -ComObject Word.Application
$template = $word.Documents.Open($templateFile)

# Replace the placeholder text in the template with the user's letter
$template.Content.Find.Execute("<letter>") | Out-Null
$template.Content.Text = $env:letter

# Save the merged letter as a new Word document
$fileName = "C:\Letter_$(Get-Date -Format yyyyMMdd-HHmmss).docx"
$template.SaveAs([ref] $fileName)
$template.Close()
$word.Quit()
