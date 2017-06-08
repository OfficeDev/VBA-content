---
title: Range.AllowEdit Property (Excel)
keywords: vbaxl10.chm144239
f1_keywords:
- vbaxl10.chm144239
ms.prod: excel
api_name:
- Excel.Range.AllowEdit
ms.assetid: 9f03054c-190f-ce3b-54db-bc6e19b7e1c6
ms.date: 06/08/2017
---


# Range.AllowEdit Property (Excel)

Returns a  **Boolean** value that indicates if the range can be edited on a protected worksheet.


## Syntax

 _expression_ . **AllowEdit**

 _expression_ A variable that represents a **Range** object.


## Example

In this example, Microsoft Excel notifies the user if cell A1 can be edited or not on a protected worksheet.


```vb
Sub UseAllowEdit() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Protect the worksheet 
 wksOne.Protect 
 
 ' Notify the user about editing cell A1. 
 If wksOne.Range("A1").AllowEdit = True Then 
 MsgBox "Cell A1 can be edited." 
 Else 
 Msgbox "Cell A1 cannot be edited." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

