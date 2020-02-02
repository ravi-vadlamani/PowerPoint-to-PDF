# Add the PowerPoint assemblies that we'll need
Add-type -AssemblyName office -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue

# Script starts PowerPoint program
$ppt = new-object -com powerpoint.application
$ppt.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Set the location where to find the PowerPoint files (also where to store the thumbnails & temp)
$pptPath = "C:\Users\Ravi\Desktop\"


# Loops through each PowerPoint file present in the directory mentioned above
Foreach($iFile in $(ls $pptPath -Filter "*.ppt")){
Set-ItemProperty ($pptPath + $iFile) -name IsReadOnly -value $false
$filename = Split-Path $iFile -leaf
$file = $filename.Split(".")[0]
$oFile = $pptPath + $file + ".pdf"

# Open the PowerPoint file
$pres = $ppt.Presentations.Open($pptPath + $iFile)

# Now saves it as PDF
$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
$pres.SaveAs($ofile,$opt)


#Closes file and deletes the temp files created
$pres.Close();

}

#Clean Up
$ppt.quit();
$ppt = $null
[gc]::Collect();
[gc]::WaitForPendingFinalizers();