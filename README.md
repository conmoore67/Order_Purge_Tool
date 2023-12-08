This PowerShell script is a comprehensive tool designed for managing and processing transportation orders. It consists of several key sections:

Module Check/Installation:
The script starts by loading the necessary Windows Forms assembly, used for creating graphical user interfaces.
A form with a progress bar is created to visually represent the installation progress of required modules.
It includes a custom function, CheckAndInstallModule, which checks for the presence of specified PowerShell modules (NuGet, SharePointPnPPowerShellOnline, Microsoft.Online.SharePoint.PowerShell, SQLSERVER) and installs them if they are not found.

Windows Forms Creation for User Input:
The script then generates a user interface for order processing. This includes a form titled 'UPT Order Purge Tool', which allows users to input and confirm a specific move or order number.
A PictureBox control displays a company logo using a Base64 encoded image.
The form has an 'OK' button that triggers a validation process to ensure the move numbers entered match. It also provides a confirmation prompt before proceeding.

SQL Query Execution, Backup, and Purge Process:
The script defines and executes a detailed SQL query to retrieve data related to the provided move number from various database tables.
Retrieved data is exported to a CSV file, which is then uploaded to a SharePoint site for backup.
Following successful backup, the script displays a progress bar while purging the corresponding data from the SQL database.
In case of a backup failure, a custom error form is shown with a bold message "Backup Unsuccessful!!" and additional instructions in standard text, enhancing the user's ability to recognize and understand the error.
Overall, this script automates critical aspects of order data management, including retrieval, backup, and purging, while providing a user-friendly interface for input and confirmation, and clear feedback in case of errors.
