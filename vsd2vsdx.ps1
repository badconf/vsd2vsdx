# Define the directory for converting VSD to VSDX files
$dir='.\'

# Create a new invisible Visio application object
$visio= New-Object -com Visio.InvisibleApp

# Get all the VSD files in the specified directory
$files = get-childitem "$dir\\*.vsd"

# Loop through each VSD file
foreach($file in $files) {
   # Print a message indicating the current file being processed
   write-host "Working on $file"

   # Open the VSD file using the Visio application
   $doc=$visio.Documents.Open($file.FullName)

   # Change the file extension from VSD to VSDX
   $vsdx=[io.path]::ChangeExtension($file,'.vsdx')

   # Save the file with the new VSDX extension
   $doc.SaveAs($vsdx)

   # Close the current document
   $doc.close();
}
