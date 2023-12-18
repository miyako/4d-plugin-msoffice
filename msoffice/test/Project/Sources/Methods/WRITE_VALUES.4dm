//%attributes = {}
var $sourcefile; $targetfile : 4D:C1709.File
$sourcefile:=File:C1566("/RESOURCES/BC22101-12-01.00 - DEV-FO-008 vI.xlsm")

$targetfile:=Folder:C1567(fk desktop folder:K87:19).file("BC22101-12-01.00 - DEV-FO-008 vI.xlsm")

$status:=Write to spreadsheet(\
$sourcefile.platformPath; \
$targetfile.platformPath; \
New object:C1471("Applicant"; New object:C1471("E8"; 100000)))

/*
set cell E8 of sheet Applicant to 100000
*/

If ($status.success)
	OPEN URL:C673($targetfile.platformPath)
End if 