# VSD to VSDX Converter

This PowerShell script converts Visio files with the `.vsd` extension to the newer `.vsdx` format. The script leverages the Visio application's COM object to open, convert, and save the files in the specified directory.

## Prerequisites

- **Microsoft Visio:** Ensure that Microsoft Visio is installed on the system where the script will be run. It uses Visio's COM object to perform the conversion.
- **PowerShell:** The script must be run using PowerShell. Ensure that you have the necessary permissions to execute scripts on your system.

## Usage

1. **Set the Directory:** Modify the `$dir` variable in the script to point to the directory containing the `.vsd` files you want to convert. By default, it's set to the current directory (`'.'`).

2. **Run the Script:** Open PowerShell and navigate to the directory where the script is located. Run the script using the following command:
   \```shell
   .\vsd2vsdx.ps1
   \```

3. **Conversion Process:** The script will process each `.vsd` file in the specified directory, convert it to `.vsdx`, and save it in the same location. Progress messages will be displayed in the console.

## Notes

- Ensure that the specified directory contains the `.vsd` files you want to convert.
- If you encounter execution policy restrictions, you may need to modify the PowerShell execution policy settings to allow the script to run.
- Always test the script on a small set of files first to ensure that it functions as expected.

## License

🔓 Licensed under [Creative Commons Public Domain (CC0)](LICENSE).
