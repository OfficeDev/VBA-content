---
title: Workbook.RemovePersonalInformation Property (Excel)
keywords: vbaxl10.chm199202
f1_keywords:
- vbaxl10.chm199202
ms.prod: excel
api_name:
- Excel.Workbook.RemovePersonalInformation
ms.assetid: f5cdc655-8ba9-6dd1-ab05-028d98c11972
ms.date: 06/08/2017
---


# Workbook.RemovePersonalInformation Property (Excel)

 **True** if personal information can be removed from the specified workbook. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **RemovePersonalInformation**

 _expression_ A variable that represents a **Workbook** object.


## Example

In this example, Microsoft Excel determines if personal information can be removed from the specified workbook and notifies the user.


```vb
Sub UsePersonalInformation() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.ActiveWorkbook 
 
 ' Determine settings and notify user. 
 If wkbOne.RemovePersonalInformation = True Then 
 MsgBox "Personal information can be removed." 
 Else 
 MsgBox "Personal information cannot be removed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

