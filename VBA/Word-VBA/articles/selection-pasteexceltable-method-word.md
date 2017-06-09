---
title: Selection.PasteExcelTable Method (Word)
keywords: vbawd10.chm158663670
f1_keywords:
- vbawd10.chm158663670
ms.prod: word
api_name:
- Word.Selection.PasteExcelTable
ms.assetid: bfa7f82c-e5c0-3919-f492-a71c9eabb919
ms.date: 06/08/2017
---


# Selection.PasteExcelTable Method (Word)

Pastes and formats a Microsoft Excel table.


## Syntax

 _expression_ . **PasteExcelTable**( **_LinkedToExcel_** , **_WordFormatting_** , **_RTF_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LinkedToExcel_|Required| **Boolean**| **True** links the pasted table to the original Excel file so that changes made to the Excel file are reflected in Microsoft Word.|
| _WordFormatting_|Required| **Boolean**| **True** formats the table using the formatting in the Word document. **False** formats the table according to the original Excel file.|
| _RTF_|Required| **Boolean**| **True** pastes the Excel table using Rich Text Format (RTF). **False** pastes the Excel table as HTML.|

## Example

This example pastes an Excel table into the active document. The parameters specify that the pasted table is linked to the Excel file, retains the original Excel formatting, and is pasted as RTF. This example assumes that the Clipboard contains an Excel table.


```vb
Sub PasteExcelFormatted() 
 Selection.PasteExcelTable _ 
 LinkedToExcel:=True, _ 
 WordFormatting:=False, _ 
 RTF:=True 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

