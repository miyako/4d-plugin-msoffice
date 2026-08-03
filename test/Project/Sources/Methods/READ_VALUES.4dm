//%attributes = {}
var $file : 4D:C1709.File
$file:=File:C1566("/RESOURCES/BC22101-12-01.00 - DEV-FO-008 vI.xlsm")

//resolve file system path such as /RESOURCES/
$file:=OB Class:C1730($file).new($file.platformPath; fk platform path:K87:2)

//all existing cell
$status:=Read from spreadsheet($file.platformPath)

If ($status.success)
	ALERT:C41(JSON Stringify:C1217($status.values; *))
End if 

//only specified cells
$status:=Read from spreadsheet($file.platformPath; \
New object:C1471("Applicant"; New collection:C1472("E8")))

If ($status.success)
	ALERT:C41(JSON Stringify:C1217($status.values; *))
End if 