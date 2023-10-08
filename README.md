# VBA Folder File Merger

This VBA script is designed to simplify the process of merging multiple files within a specified folder into a single consolidated file. Whether you're dealing with Excel workbooks, text files, or any other type of compatible files, this script can help streamline the process.

## Getting Started

### Copy the VBA code from the .txt file

Start by either downloading the file the contains the code to your local machine or simply copy and paste it into the VBA editor.

### Open Excel

Ensure that you have Microsoft Excel installed on your computer. This script is designed to run within Excel.

### Enable Macros

To run VBA scripts, you need to enable macros in Excel. Go to Excel Options > Trust Center > Trust Center Settings > Macro Settings, and select "Enable all macros" or "Enable macros with notification," depending on your security preferences.

### Open the VBA Editor

Press `Alt` + `F11` to open the Visual Basic for Applications (VBA) editor in Excel.

### Import the Script

In the VBA editor, go to File > Import File, and select the VBA script file you downloaded from this repository.

### Configure Parameters

Inside the VBA script, you can configure the folder path where your files are located, the file extension(s) to include (it only checks for excel files by default), and the name of the merged file.

### Run the Script

Close the VBA editor and return to your Excel workbook. You can then run the script by going to the Developer tab (if not visible, enable it in Excel options) and click on "Run." Alternatively, you can assign the script to a button or keyboard shortcut.

### Review the Merged File

Once the script has run successfully, you'll find the merged file in the specified folder.

## Configuration

You can customize the behavior of this script by modifying the following parameters:

- `MyPath`: Set the folder path where your files are located.
- `MasterWorkbook.SaveAs`: Specify the filepath and name of the merged file.

## Compatibility

This script is compatible with Microsoft Excel.

## Contributing

If you'd like to contribute to this project or suggest improvements, feel free to open an issue or create a pull request.

## Acknowledgments

- This script was developed to simplify the process of merging files for various tasks, saving time and effort in data consolidation.
- Special thanks to the VBA community for their valuable insights and contributions.
