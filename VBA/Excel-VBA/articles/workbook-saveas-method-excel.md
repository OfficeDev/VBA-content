---
title: Workbook.SaveAs Method (Excel)
keywords: vbaxl10.chm199200
f1_keywords:
- vbaxl10.chm199200
ms.prod: excel
api_name:
- Excel.Workbook.SaveAs
ms.assetid: fbc3ce55-27a3-aa07-3fdb-77b0d611e394
ms.date: 06/08/2017
---


# Workbook.SaveAs Method (Excel)

Saves changes to the workbook in a different file.


## Syntax

 _expression_ . **SaveAs**( **_FileName_** , **_FileFormat_** , **_Password_** , **_WriteResPassword_** , **_ReadOnlyRecommended_** , **_CreateBackup_** , **_AccessMode_** , **_ConflictResolution_** , **_AddToMru_** , **_TextCodepage_** , **_TextVisualLayout_** , **_Local_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Optional| **Variant**|A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.|
| _FileFormat_|Optional| **Variant**|The file format to use when you save the file. For a list of valid choices, see the  **[XlFileFormat](xlfileformat-enumeration-excel.md)** enumeration. For an existing file, the default format is the last file format specified; for a new file, the default is the format of the version of Excel being used.|
| _Password_|Optional| **Variant**|A case-sensitive string (no more than 15 characters) that indicates the protection password to be given to the file.|
| _WriteResPassword_|Optional| **Variant**|A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only.|
| _ReadOnlyRecommended_|Optional| **Variant**| **True** to display a message when the file is opened, recommending that the file be opened as read-only.|
| _CreateBackup_|Optional| **Variant**| **True** to create a backup file.|
| _AccessMode_|Optional| **[XlSaveAsAccessMode](xlsaveasaccessmode-enumeration-excel.md)**|The access mode for the workbook.|
| _ConflictResolution_|Optional| **[XlSaveConflictResolution](xlsaveconflictresolution-enumeration-excel.md)**|An  **[XlSaveConflictResolution](xlsaveconflictresolution-enumeration-excel.md)** value that determines how the method resolves a conflict while saving the workbook. If set to **xlUserResolution** , the conflict-resolution dialog box is displayed. If set to **xlLocalSessionChanges** , the local user's changes are automatically accepted. If set to **xlOtherSessionChanges** , the changes from other sessions are automatically accepted instead of the local user's changes. If this argument is omitted, the conflict-resolution dialog box is displayed.|
| _AddToMru_|Optional| **Variant**| **True** to add this workbook to the list of recently used files. The default value is **False** .|
| _TextCodepage_|Optional| **Variant**|Ignored for all languages in Microsoft Excel.<table><tr><th>**Note**</th></tr><tr><td>When Excel saves a workbook to one of the CSV or text formats, which are specified by using the FileFormat parameter, it uses the code page that corresponds to the language for the system locale in use on the current computer. This system setting is available in the  **Control Panel**, by clicking  **Region and Language**, clicking the  **Location** tab, under **Current location**.</td></tr></table>|
| _TextVisualLayout_|Optional| **Variant**|Ignored for all languages in Microsoft Excel.<table><tr><th>**Note**</th></tr><tr><td>When Excel saves a workbook to one of the CSV or text formats, which are specified by using the FileFormat parameter, it saves these formats in logical layout. If left-to-right (LTR) text is embedded within right-to-left (RTL) text in the file, or vice versa, logical layout saves the contents of the file in the correct reading order for all languages in the file without regard to direction. When an application opens the file, each run of LTR or RTL characters are rendered in the correct direction according to the character value ranges within the code page. (Unless an application that is designed to display the exact memory layout of the file, such as a debugger or editor, is used to open the file.)</td></tr></table>|
| _Local_|Optional| **Variant**| **True** saves files against the language of Microsoft Excel (including control panel settings). **False** (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically US English unless the VBA project where **Workbooks.Open** is run from is an old internationalized XL5/95 VBA project).|

## Remarks

Use strong passwords that combine upper- and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Use a strong password that you can remember so that you don't have to write it down


## Example

This example creates a new workbook, prompts the user for a file name, and then saves the workbook.


```vb
Set NewBook = Workbooks.Add 
Do 
    fName = Application.GetSaveAsFilename 
Loop Until fName <> False 
NewBook.SaveAs Filename:=fName
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

