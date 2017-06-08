---
title: Workbook.UpdateLinks Property (Excel)
keywords: vbaxl10.chm199197
f1_keywords:
- vbaxl10.chm199197
ms.prod: excel
api_name:
- Excel.Workbook.UpdateLinks
ms.assetid: c8d374d7-0b32-eb32-fa29-ab496d6786e7
ms.date: 06/08/2017
---


# Workbook.UpdateLinks Property (Excel)

Returns or sets an  **[XlUpdateLink](xlupdatelinks-enumeration-excel.md)** constant indicating a workbook's setting for updating embedded OLE links. Read/write.


## Syntax

 _expression_ . **UpdateLinks**

 _expression_ A variable that represents a **Workbook** object.


## Remarks





| **XlUpdateLinks** can be one of these **XlUpdateLinks** constants.|
| **xlUpdateLinksAlways** Embedded OLE links are always updated for the specified workbook.|
| **xlUpdateLinksNever** Embedded OLE links are never updated for the specified workbook.|
| **xlUpdateLinksUserSetting** Embedded OLE links are updated according to the user's settings for the specified workbook.|

## Example

In this example, Microsoft Excel determines the setting for updating links and notifies the user.


```vb
Sub UseUpdateLinks() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.Workbooks(1) 
 
 Select Case wkbOne.UpdateLinks 
 Case xlUpdateLinksAlways 
 MsgBox "Links will always be updated " &; _ 
 "for the specified workbook." 
 Case xlUpdateLinksNever 
 MsgBox "Links will never be updated " &; _ 
 "for the specified workbook." 
 Case xlUpdateLinksUserSetting 
 MsgBox "Links will update according " &; _ 
 "to user settting for the specified workbook." 
 End Select 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

