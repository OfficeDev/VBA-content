---
title: Worksheet.FollowHyperlink Event (Excel)
keywords: vbaxl10.chm502080
f1_keywords:
- vbaxl10.chm502080
ms.prod: excel
api_name:
- Excel.Worksheet.FollowHyperlink
ms.assetid: c63eec19-008e-bfb5-1357-3d02426c1bab
ms.date: 06/08/2017
---


# Worksheet.FollowHyperlink Event (Excel)

Occurs when you click any hyperlink on a worksheet. For application- and workbook-level events, see the  **[SheetFollowHyperlink](application-sheetfollowhyperlink-event-excel.md)** event and **[SheetFollowHyperlink](workbook-sheetfollowhyperlink-event-excel.md)** event.


## Syntax

 _expression_ . **FollowHyperlink**( **_Target_** )

 _expression_ An expression that returns a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **Hyperlink**|A  **[Hyperlink](hyperlink-object-excel.md)** object that represents the destination of the hyperlink.|

## Example

This example keeps a list, or history, of all the links that have been visited from the active worksheet.


```vb
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink) 
    With UserForm1 
        .ListBox1.AddItem Target.Address 
        .Show 
    End With 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

