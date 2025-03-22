# LetterServer

A PowerShell-based web server that provides an interface for writing letters and merging them with Word templates.

## Overview

LetterServer creates a simple web server on a specified port and provides an HTML interface for users to write letters. The script can merge the letter content with a specified Word template and save the resulting document.

## Features

- **Simple Web Interface**: Clean, responsive interface for writing letters
- **Word Template Integration**: Merges letter content with Word templates
- **Customizable Port**: Run the server on any available port
- **Automatic File Saving**: Saves merged letters with timestamps
- **Error Handling**: Robust error handling and user feedback

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- Microsoft Word installed
- Required PowerShell modules:
  - WebAdministration
  - Microsoft.Office.Interop.Word

## Installation

1. Clone this repository or download the files to your local machine
2. Ensure you have the required PowerShell modules installed:
   ```powershell
   Install-Module -Name WebAdministration -Force
   ```
3. Make sure Microsoft Word is installed on your system

## Usage

1. Run the PowerShell script:
   ```powershell
   .\letter-server.ps1
   ```

2. The script will prompt you for:
   - Port number for the web server (default: 8080)
   - Physical path for the website files (default: script directory\WebSite)
   - Path to a Word template file (default: script directory\LetterTemplate.dotx)

3. If the specified Word template doesn't exist, the script will create a simple template for you

4. Access the letter writing interface by opening a web browser and navigating to:
   ```
   http://localhost:8080/
   ```
   (Replace 8080 with your chosen port number)

5. Write your letter in the text area and click "Submit Letter"

6. The letter will be merged with the Word template and saved as a Word document in the script directory with a timestamp in the filename

## File Structure

- `letter-server.ps1`: Main PowerShell script that runs the web server
- `index.html`: HTML template for the letter writing interface
- `letter.html`: HTML template for displaying the submitted letter

## Customization

### Word Template

You can create your own Word template with the placeholder `<letter>` where you want the letter content to be inserted. If the placeholder isn't found, the letter content will be appended to the end of the document.

### HTML Templates

You can modify the HTML files to customize the appearance of the web interface. The script copies these files to the website directory when it starts.

## Troubleshooting

- **Module Not Found**: If you receive errors about missing modules, make sure you've installed the required PowerShell modules and have Microsoft Word installed
- **Port Already in Use**: If the specified port is already in use, choose a different port number
- **Permission Issues**: Make sure you're running PowerShell with appropriate permissions to create websites and access the file system

## License

This project is licensed under the MIT License - see the LICENSE file for details.
