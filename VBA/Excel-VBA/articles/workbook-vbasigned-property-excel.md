---
title: Workbook.VBASigned Property (Excel)
keywords: vbaxl10.chm199195
f1_keywords:
- vbaxl10.chm199195
ms.prod: excel
api_name:
- Excel.Workbook.VBASigned
ms.assetid: 6e93161c-2fa4-1064-9b5d-a8eb96ad2bea
ms.date: 06/08/2017
---


# Workbook.VBASigned Property (Excel)

 **True** if the Visual Basic for Applications project for the specified workbook has been digitally signed. Read-only **Boolean** .


## Syntax

 _expression_ . **VBASigned**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example loads a workbook named "mybook.xls" and then tests to see whether its Visual Basic for Applications project has a digital signature. If there's no digital signature, the example displays a warning message.


```vb
Workbooks.Open FileName:="c:\My Documents\mybook.xls", _ 
 ReadOnly:=False 
If Workbook.VBASigned = False Then 
 MsgBox "Warning! The project " _ &; 
 "has not been digitally signed." _ &; 
 , vbCritical, "Digital Signature Warning" 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

