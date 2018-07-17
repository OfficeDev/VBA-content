---
title: SmartDocument.SolutionID Property (Office)
keywords: vbaof11.chm262001
f1_keywords:
- vbaof11.chm262001
ms.prod: office
api_name:
- Office.SmartDocument.SolutionID
ms.assetid: e1eea0af-d951-a316-4b58-a65ffd280c89
ms.date: 06/08/2017
---


# SmartDocument.SolutionID Property (Office)

Gets or sets the ID, often a globally unique identifier (GUID), which identifies the XML expansion pack attached to the active document in Microsoft Word or workbook in Microsoft Excel. Read/write.


## Syntax

 _expression_. **SolutionID**

 _expression_ A variable that represents a **SmartDocument** object.


## Remarks

The  **SolutionID** property returns an empty string or "None" when no XML expansion pack is attached to the active document.

Provide appropriate values for the  **SolutionID** and **SolutionURL** properties to attach an available XML expansion pack to the active document to transform it into a smart document without using the **PickSolution** method. Set the **SolutionID** and **SolutionUrl** properties to empty strings to remove the attached XML expansion pack.


## Example

The following example determines whether an XML expansion pack is attached to the active Excel workbook by checking the  **SolutionID** property.


```
 Dim objSmartDoc As Office.SmartDocument 
 Set objSmartDoc = ActiveWorkbook.SmartDocument 
 If objSmartDoc.SolutionID = "None" Or objSmartDoc.SolutionID = "" Then 
 MsgBox "No XML expansion pack attached." 
 Else 
 MsgBox "Smart document Solution ID: " &amp; _ 
 objSmartDoc.SolutionID 
 End If 
 Set objSmartDoc = Nothing 

```


## See also


#### Concepts


[SmartDocument Object](smartdocument-object-office.md)
#### Other resources


[SmartDocument Object Members](smartdocument-members-office.md)

