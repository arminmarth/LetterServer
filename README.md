# LetterServer

LetterServer is a Powershell script that creates a web server on a specified port and provides a HTML interface for users to write a letter. The script can also merge the letter with a specified Word template and save the resulting document.

## Usage

1. Install the `WebAdministration` and `Microsoft.Office.Interop.Word` modules on your system.
2. Create two new files named `index.html` and `letter.html` in the same directory as the Powershell script. These files should contain the HTML code for the default page and the letter display page, respectively.
3. Update the `$templateFile`

