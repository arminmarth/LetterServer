# LetterServer

LetterServer is a Powershell script that creates a web server on a specified port and provides a HTML interface for users to write a letter. The script can also merge the letter with a specified Word template and save the resulting document.

## Usage

1. Install the `WebAdministration` and `Microsoft.Office.Interop.Word` modules on your system.
2. Update the `$templateFile` variable in the script with the path to your Word template file.
3. Run the Powershell script: `.\letter_server.ps1`
4. Follow the prompts to specify the port number, physical path, and Word template file to use for the website and letter generation.
5. Access the website at http://localhost:<port_number> (replace `<port_number>` with the port number you specified in step 4).
6. Write your letter using the HTML interface on the website.
7. Submit the letter to view the merged document and save it as a new Word file.
