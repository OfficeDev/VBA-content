---
title: Worksheet.SaveAs Method (Excel)
keywords: vbaxl10.chm175151
f1_keywords:
- vbaxl10.chm175151
ms.prod: excel
api_name:
- Excel.Worksheet.SaveAs
ms.assetid: 2c20ccd0-c4b8-599f-3923-a432caeb6b91
ms.date: 06/08/2017
---


# Worksheet.SaveAs Method (Excel)

Saves changes to the chart or worksheet in a different file.


## Syntax

 _expression_ . **SaveAs**( **_FileName_** , **_FileFormat_** , **_Password_** , **_WriteResPassword_** , **_ReadOnlyRecommended_** , **_CreateBackup_** , **_AddToMru_** , **_TextCodepage_** , **_TextVisualLayout_** , **_Local_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**| **Variant** . A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.|
| _FileFormat_|Optional| **Variant**|The file format to use when you save the file. For a list of valid choices, see the  **[XlFileFormat](xlfileformat-enumeration-excel.md)** enumeration. For an existing file, the default format is the last file format specified; for a new file, the default is the format of the version of Excel being used.|
| _Password_|Optional| **Variant**|A case-sensitive string (no more than 15 characters) that indicates the protection password to be given to the file.|
| _WriteResPassword_|Optional| **Variant**|A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only.|
| _ReadOnlyRecommended_|Optional| **Variant**| **True** to display a message when the file is opened, recommending that the file be opened as read-only.|
| _CreateBackup_|Optional| **Variant**| **True** to create a backup file.|
| _AddToMru_|Optional| **Variant**| **True** to add this workbook to the list of recently used files. The default value is **False** .|
| _TextCodepage_|Optional| **Variant**|Not used in U.S. English Microsoft Excel.|
| _TextVisualLayout_|Optional| **Variant**|Not used in U.S. English Microsoft Excel.|
| _Local_|Optional| **Variant**| **True** saves files against the language of Microsoft Excel (including control panel settings). **False** (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically US English unless the VBA project where Workbooks.Open is run from is an old internationalized XL5/95 VBA project).|

## Remarks

Use strong passwords that combine upper- and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27. Use a strong password that you can remember so that you don't have to write it down


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

