---
title: Worksheet.AutoFilterMode Property (Excel)
keywords: vbaxl10.chm175075
f1_keywords:
- vbaxl10.chm175075
ms.prod: excel
api_name:
- Excel.Worksheet.AutoFilterMode
ms.assetid: 63f33ea5-c9a5-0096-0191-1590cda9d0e1
ms.date: 06/08/2017
---


# Worksheet.AutoFilterMode Property (Excel)

 **True** if the AutoFilter drop-down arrows are currently displayed on the sheet. This property is independent of the **FilterMode** property. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFilterMode**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

This property returns  **True** if the drop-down arrows are currently displayed. You can set this property to **False** to remove the arrows, but you cannot set it to **True** . Use the **[AutoFilter](worksheet-autofilter-property-excel.md)** method to filter a list and display the drop-down arrows.


## Example

This example displays the current status of the  **AutoFilterMode** property on Sheet1.


```vb
If Worksheets("Sheet1").AutoFilterMode Then 
 isOn = "On" 
Else 
 isOn = "Off" 
End If 
MsgBox "AutoFilterMode is " &; isOn
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

