---
title: Workbook.SheetFollowHyperlink Event (Excel)
keywords: vbaxl10.chm503092
f1_keywords:
- vbaxl10.chm503092
ms.prod: excel
api_name:
- Excel.Workbook.SheetFollowHyperlink
ms.assetid: be29df8c-4e8e-f719-ae1d-f91a11b89491
ms.date: 06/08/2017
---


# Workbook.SheetFollowHyperlink Event (Excel)

Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the  **[FollowHyperlink](worksheet-followhyperlink-event-excel.md)** event.


## Syntax

 _expression_ . **SheetFollowHyperlink**( **_Sh_** , **_Target_** )

 _expression_ An expression that returns a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The  **[Worksheet](worksheet-object-excel.md)** object that contains the hyperlink.|
| _Target_|Required| **Hyperlink**|The  **[Hyperlink](hyperlink-object-excel.md)** object that represents the destination of the hyperlink.|

## Example

This example keeps a list, or history, of all the hyperlinks in the current workbook that have been clicked, plus the names of the worksheets that contain these hyperlinks.


```vb
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh as Object, _ 
 ByVal Target As Hyperlink) 
 UserForm1.ListBox1.AddItem Sh.Name &; ":" &; Target.Address 
 UserForm1.Show 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

