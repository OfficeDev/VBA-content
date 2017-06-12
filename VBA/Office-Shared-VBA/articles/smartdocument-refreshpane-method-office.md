---
title: SmartDocument.RefreshPane Method (Office)
keywords: vbaof11.chm262004
f1_keywords:
- vbaof11.chm262004
ms.prod: office
api_name:
- Office.SmartDocument.RefreshPane
ms.assetid: c37de2c2-f24a-0db2-fda8-cfe7d0b464fb
ms.date: 06/08/2017
---


# SmartDocument.RefreshPane Method (Office)

Refreshes the  **Document Actions** task pane for the active document in Microsoft Word or a workbook in Microsoft Excel.


## Syntax

 _expression_. **RefreshPane**

 _expression_ A variable that represents a **SmartDocument** object.


## Remarks

The  **RefreshPane** method raises an error if the active document does not have an XML expansion pack attached.


## Example

The following example determines whether the active Excel workbook has an XML expansion pack attached. If so, it refreshes the smart document's  **Document Actions** task pane.


```
 Dim objSmartDoc As Office.SmartDocument 
 Set objSmartDoc = ActiveWorkbook.SmartDocument 
 If objSmartDoc.SolutionID > "None" Then 
 objSmartDoc.RefreshPane 
 Else 
 MsgBox "No XML expansion pack attached." 
 End If 

```


## See also


#### Concepts


[SmartDocument Object](smartdocument-object-office.md)
#### Other resources


[SmartDocument Object Members](smartdocument-members-office.md)

