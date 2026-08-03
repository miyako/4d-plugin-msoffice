//%attributes = {}
var $file : 4D:C1709.File
$file:=File:C1566("/RESOURCES/BC22101-12-01.00 - DEV-FO-008 vI.xlsm")

var $XLSX : Blob  //declare it, or else it could be 4D.Blob (object) which is not compatible with the plugin
$XLSX:=$file.getContent()

$status:=Verify office document($XLSX)

If ($status.success) & ($status.format="Zip")
	
	$file:=Folder:C1567(fk desktop folder:K87:19).file("encrypted.xlsm")
	
	$status:=Encode office document($XLSX; $file.platformPath; "PASSWORD"; MSOFFICE Encryption AES256)
	
	If ($status.success)
		
		$XLSX:=$file.getContent()
		
		$file:=Folder:C1567(fk desktop folder:K87:19).file("decrypted.xlsm")
		
		$status:=Decode office document($XLSX; Folder:C1567(fk desktop folder:K87:19).file("decrypted.xlsm").platformPath; "PASSWORD")
		
	End if 
	
End if 