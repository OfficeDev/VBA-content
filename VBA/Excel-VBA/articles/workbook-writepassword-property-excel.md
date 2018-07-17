---
title: Workbook.WritePassword Property (Excel)
keywords: vbaxl10.chm199210
f1_keywords:
- vbaxl10.chm199210
ms.prod: excel
api_name:
- Excel.Workbook.WritePassword
ms.assetid: ac89063e-6ef5-f7c5-abb0-4e6ef1c5fd05
ms.date: 06/08/2017
---


# Workbook.WritePassword Property (Excel)

Returns or sets a  **String** for the write password of a workbook. Read/write.


## Syntax

 _expression_ . **WritePassword**

 _expression_ A variable that represents a **Workbook** object.


## Example

In this example, if the active workbook is not protected against saving changes, Microsoft Excel sets the password to a string as the write password for the active workbook.


```vb
Sub UseWritePassword() 
 
 Dim strPassword As String 
 
 strPassword = InputBox ("Enter the password") 
 
 ' Set password to a string if allowed. 
 If ActiveWorkbook.WriteReserved = False Then 
 ActiveWorkbook.WritePassword = strPassword 
 End If 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

