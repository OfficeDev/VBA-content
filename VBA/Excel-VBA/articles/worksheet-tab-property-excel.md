---
title: Worksheet.Tab Property (Excel)
keywords: vbaxl10.chm175149
f1_keywords:
- vbaxl10.chm175149
ms.prod: excel
api_name:
- Excel.Worksheet.Tab
ms.assetid: 386edcb0-868e-3f24-b4f0-8e52b9fcffcb
ms.date: 06/08/2017
---


# Worksheet.Tab Property (Excel)

Returns a  **[Tab](tab-object-excel.md)** object for a worksheet.


## Syntax

 _expression_ . **Tab**

 _expression_ A variable that represents a **Worksheet** object.


## Example

In this example, Microsoft Excel determines if the worksheet's first tab color index is set to none and notifies the user.


```vb
Sub CheckTab() 
 
 ' Determine if color index of 1st tab is set to none. 
 If Worksheets(1).Tab.ColorIndex = xlColorIndexNone Then 
 MsgBox "The color index is set to none for the 1st " &; _ 
 "worksheet tab." 
 Else 
 MsgBox "The color index for the tab of the 1st worksheet " &; _ 
 "is not set none." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

