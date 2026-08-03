# 4d-plugin-msoffice

The MSOFFICE plugin reads and writes cell values in Excel spreadsheets, inspects whether a document is a modern (`Zip`-based, OOXML) or legacy (`Cfb`/OLE) Office file, and encrypts/decrypts Office documents with a password using AES. Spreadsheet reading and writing is implemented on top of [OpenXLSX](https://github.com/troldal/OpenXLSX); document encoding/decoding and format detection work directly on the raw file bytes you pass in as a `Blob`, not on a file path. Every command returns a 4D `Object` — there is no separate error-signaling mechanism, so `.success` (and, for some commands, `.error`) is always how you tell whether the call did what you asked.

## Summary table

| Command | Returns | Purpose |
|---|---|---|
| [`Write to spreadsheet`](#write-to-spreadsheet) | `Object` | Copy a spreadsheet file and set specific cell values in the copy |
| [`Read from spreadsheet`](#read-from-spreadsheet) | `Object` | Read cell values, sheet names, and document properties from a spreadsheet file |
| [`Verify office document`](#verify-office-document) | `Object` | Detect whether raw document bytes are a modern (Zip/OOXML) or legacy (Cfb/OLE) Office file |
| [`Decode office document`](#decode-office-document) | `Object` | Decrypt password-protected document bytes to a plain file on disk |
| [`Encode office document`](#encode-office-document) | `Object` | Encrypt document bytes with a password (AES-128 or AES-256) to a file on disk |

**Platforms:** macOS, Windows

---

## Requirements & platform notes

- All five commands are declared `threadSafe` in the plugin manifest — they can be called from any 4D process or worker, including preemptive ones.
- `Write to spreadsheet` and `Read from spreadsheet` operate on **file paths** (`platformPath`); `Verify office document`, `Decode office document`, and `Encode office document` operate on **raw file bytes** (`Blob`) you've already loaded into memory (e.g. via `File.getContent()`), not on a path.
- **`Write to spreadsheet`'s 4th parameter (`password`) is accepted but currently has no effect.** The manifest declares it and the plugin will accept a call that passes it, but the active code path never reads it — the output file is written unencrypted regardless of what you pass here. If you need password protection on the result, write the file first and then encrypt it separately with [`Encode office document`](#encode-office-document).
- **`Read from spreadsheet`'s 3rd parameter is declared in the manifest but never read by the plugin.** Passing a value there has no effect; it's safe to omit.
- `Write to spreadsheet` silently does nothing (returns `{success: false}` with no error message) if `source` and `target` are the same path — there's no in-place edit mode.
- `Decode office document` and `Encode office document` don't populate an `.error` property on failure — you only get `.success = false`, with no further diagnostic. `Write to spreadsheet` and `Read from spreadsheet` are more informative on failure (see each command's Description below).
- "Zip" vs. "Cfb" (used by `Verify office document`) are the two container formats Office documents come in: modern `.xlsx`/`.docx`/`.xlsm` files are Zip archives (OOXML); legacy `.xls`/`.doc` files are OLE Compound File Binary. `Verify office document` tells you which one you're holding before you try to decode or write to it.

---

## Write to spreadsheet

### Syntax
```4d
Write to spreadsheet ( source ; target ; references ; password ) → Object
```

| Parameter | Type | Description |
|---|---|---|
| `source` | Text | Path of the existing spreadsheet file to copy from. |
| `target` | Text | Path to write the modified copy to. Must be different from `source` — see below. |
| `references` | Object | Maps sheet names to an object of cell references and the values to set, e.g. `{"Sheet1": {"A1": "hello", "B2": 42}}`. |
| `password` | Text | **Currently ignored.** Accepted for compatibility; passing a value has no effect on the output file. |
| Result | Object | `{success: Boolean; temporaryFilePath: Text; error: Text}` — see Description. |

### Description
The plugin copies `source` to `target` (via the underlying platform's native file-copy command), applies each cell in `references` to a working copy of `source`, then merges that back into `target`, preserving macros and other parts of the original file (VBA projects and ActiveX parts are copied through unmodified from `source`; a handful of internal customXml/metadata parts are excluded from that pass-through).

- Under each sheet name in `references`, you can set `String`, `Boolean`, `Null` (clears the cell), or numeric (`Int`, `Int64`, `Double`) values — the plugin dispatches on the JSON value's type. Any other JSON type for a cell is silently skipped.
- If `source == target`, or `references` is missing/empty, or a sheet name in `references` doesn't exist in the workbook, the corresponding writes are skipped — you get `{success: false}` with **no error message**, since no exception was thrown.
- If an internal error does occur partway through (e.g. the document can't be opened, or saving fails), you get `{success: false, error: "<message>"}`. Note that `temporaryFilePath` may still be present in the result even when `success` is `false`, since it's set before the save/merge steps that can fail.
- On success, `temporaryFilePath` points at an intermediate file the plugin created during the merge — informational only, not something you need to clean up or rely on.

### Example
From the plugin's own sample method (`WRITE_VALUES.4dm`):
```4d
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
```

Setting several cells across two sheets in one call:
```4d
var $refs : Object
$refs:=New object:C1471(\
"Applicant"; New object:C1471("E8"; 100000; "E9"; "Approved"); \
"Summary"; New object:C1471("B2"; True))

$status:=Write to spreadsheet($source.platformPath; $target.platformPath; $refs)

If (Not($status.success))
	ALERT:C41("Write failed"+(Length:C16($status.error) # 0 ? ": "+$status.error : ""))
End if 
```

---

## Read from spreadsheet

### Syntax
```4d
Read from spreadsheet ( path ; references ) → Object
```

| Parameter | Type | Description |
|---|---|---|
| `path` | Text | Path of the spreadsheet file to read. |
| `references` | Object | Optional. Maps sheet names to a collection of cell references to read, e.g. `{"Applicant": ["E8", "E9"]}`. If omitted, every cell of every sheet is returned instead. |
| Result | Object | `{success: Boolean; core_property: Object; values: Object; sheets: Collection}` — see Description. |

### Description
- `core_property` is always populated on success, with these Text fields: `title`, `subject`, `creator`, `keywords`, `description`, `category`, `last_modified_by`, `created`, `modified`, `last_printed`, `version` — read from the document's OOXML core properties.
- `sheets` is a Collection of every worksheet name in the workbook, regardless of whether `references` was passed.
- **If `references` is omitted, `values` contains every cell of every sheet.** For a large workbook this reads and returns the entire sheet content, not just a summary — pass `references` to limit the read to specific cells.
- If `references` is passed, `values` only contains the cells you asked for, grouped by sheet name; a sheet name in `references` that doesn't exist in the workbook is skipped.
- Each cell entry under `values` is an object keyed by cell reference (e.g. `"E8"`), containing `data_type` (one of `"empty"`, `"boolean"`, `"number"`, `"error"`, `"string"`) and either `value` (Boolean/Number) or `stringValue` (Text/error).
- **`success` is set to `true` as soon as the file opens successfully — before properties or cell values are read.** If something fails partway through the property or cell read, you'll still get `success: true` back, just with incomplete `core_property`/`values`. Treat a mostly-empty result with `success: true` as a signal something went wrong, since the flag alone won't tell you.
- On failure to open the file at all, you get `{success: false}` with no further detail.

### Example
From the plugin's own sample method (`READ_VALUES.4dm`):
```4d
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
```

Listing sheet names and reading one cell from each:
```4d
$status:=Read from spreadsheet($file.platformPath)

If ($status.success)
	For each ($sheet; $status.sheets)
		If (OB Instance of:C1731($status.values[$sheet]))
			$firstCell:=OB Keys:C1728($status.values[$sheet])[0]
			ALERT:C41($sheet+": "+String:C10($status.values[$sheet][$firstCell]))
		End if 
	End for each 
End if 
```

---

## Verify office document

### Syntax
```4d
Verify office document ( data ) → Object
```

| Parameter | Type | Description |
|---|---|---|
| `data` | Blob | Raw bytes of the document to inspect. |
| Result | Object | `{success: Boolean; format: Text}` — `format` is `"Zip"`, `"Cfb"`, or `"Unknown"`. |

### Description
Inspects the raw bytes and reports which container format they're in, without needing a file on disk. `success` is `true` only when the format is recognized (`"Zip"` or `"Cfb"`); for anything else — including corrupt data or a completely unrelated file — you get `success: false, format: "Unknown"`.

Use this before [`Decode office document`](#decode-office-document) or [`Write to spreadsheet`](#write-to-spreadsheet) to confirm you actually have an Office document before spending time on it.

### Example
From the plugin's own sample method (`PROTECT.4dm`):
```4d
//%attributes = {}
var $file : 4D:C1709.File
$file:=File:C1566("/RESOURCES/BC22101-12-01.00 - DEV-FO-008 vI.xlsm")

var $XLSX : Blob  //declare it, or else it could be 4D.Blob (object) which is not compatible with the plugin
$XLSX:=$file.getContent()

$status:=Verify office document($XLSX)

If ($status.success) & ($status.format="Zip")
	// proceed — this is a modern OOXML file
End if 
```

---

## Decode office document

### Syntax
```4d
Decode office document ( data ; path ; password ) → Object
```

| Parameter | Type | Description |
|---|---|---|
| `data` | Blob | Raw bytes of the password-protected document. |
| `path` | Text | Path to write the decrypted document to. |
| `password` | Text | The password the document was encrypted with. |
| Result | Object | `{success: Boolean}`. |

### Description
Decrypts `data` using `password` and writes the plain (unencrypted) document to `path`. There's no `.error` field on failure — a wrong password, corrupted input, or write failure all just come back as `success: false`. Check the result with [`Verify office document`](#verify-office-document) beforehand if you're not certain `data` is actually an encrypted Office document.

### Example
From the plugin's own sample method (`PROTECT.4dm`), decoding a file that was just encoded with [`Encode office document`](#encode-office-document):
```4d
$status:=Encode office document($XLSX; $file.platformPath; "PASSWORD"; MSOFFICE Encryption AES256)

If ($status.success)
	
	$XLSX:=$file.getContent()
	
	$file:=Folder:C1567(fk desktop folder:K87:19).file("decrypted.xlsm")
	
	$status:=Decode office document($XLSX; Folder:C1567(fk desktop folder:K87:19).file("decrypted.xlsm").platformPath; "PASSWORD")
	
End if 
```

Guarding against a wrong password:
```4d
$status:=Decode office document($encryptedBytes; $outPath; $password)

If (Not($status.success))
	ALERT:C41("Could not decode the document — check the password.")
End if 
```

---

## Encode office document

### Syntax
```4d
Encode office document ( data ; path ; password ; mode ) → Object
```

| Parameter | Type | Description |
|---|---|---|
| `data` | Blob | Raw bytes of the document to encrypt. |
| `path` | Text | Path to write the encrypted document to. |
| `password` | Text | Password to protect the document with. |
| `mode` | Longint | AES key size. `MSOFFICE Encryption AES256`, seen in the sample code, is one of the two named constants the plugin exposes; the other corresponds to AES-128. Check the plugin's constants list in the Explorer for the exact name if you need AES-128. |
| Result | Object | `{success: Boolean}`. |

### Description
Encrypts `data` with `password` at the given AES key size and writes the result to `path`. Like `Decode office document`, there's no `.error` field — a failure (bad input bytes, an unwritable `path`, etc.) just comes back as `success: false`.

### Example
From the plugin's own sample method (`PROTECT.4dm`):
```4d
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
		// $file now holds the encrypted document
	End if 
	
End if 
```

---

## Error handling & troubleshooting

- **`Write to spreadsheet` no-ops silently if `source` equals `target`.** There's no in-place edit mode — always write to a different path, even a temporary one, and move it afterward if you need to replace the original.
- **A `Write to spreadsheet` failure may still return a `temporaryFilePath`.** Don't treat the presence of that field as a sign the call succeeded — always check `.success`.
- **`password` on `Write to spreadsheet` does nothing in the current build.** If your output needs to be password-protected, chain a call to [`Encode office document`](#encode-office-document) afterward.
- **The 3rd parameter of `Read from spreadsheet` is a no-op** — it's declared in the manifest but never read. Omit it.
- **`Read from spreadsheet`'s `success` can be `true` with an incomplete result.** It's set right after the file opens, before properties and cell values are read, and isn't rolled back if something fails afterward — a `true` result with missing `core_property` fields or an empty `values` object is worth treating as a partial failure.
- **`Decode office document` and `Encode office document` give you no diagnostic on failure** — just `success: false`. Use [`Verify office document`](#verify-office-document) first to rule out "this isn't actually an Office document" as the cause.
- **A sheet name in `references` that doesn't exist in the workbook is silently skipped**, for both `Write to spreadsheet` and `Read from spreadsheet` — double-check sheet name spelling/case if a cell you expect isn't showing up.
- **`references` must be a real `Object`, not a JSON-encoded Text.** Declare Blob parameters explicitly as `Blob` (not `4D.Blob`), as noted in the sample code — an object-typed Blob isn't accepted by the plugin.

---

## Quick reference

```4d
// Detect format, then encrypt
$data:=$file.getContent()
$status:=Verify office document($data)
If ($status.success) & ($status.format="Zip")
	$status:=Encode office document($data; $encPath; $pw; MSOFFICE Encryption AES256)
End if 

// Decrypt back
$status:=Decode office document($file.getContent(); $decPath; $pw)

// Write specific cells
$status:=Write to spreadsheet($srcPath; $dstPath; New object:C1471("Sheet1"; New object:C1471("A1"; 42)))

// Read specific cells
$status:=Read from spreadsheet($path; New object:C1471("Sheet1"; New collection:C1472("A1"; "B2")))

// Read everything
$status:=Read from spreadsheet($path)
$allValues:=$status.values
$sheetNames:=$status.sheets
```
