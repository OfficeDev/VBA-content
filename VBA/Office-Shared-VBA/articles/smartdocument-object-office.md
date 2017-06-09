---
title: SmartDocument Object (Office)
keywords: vbaof11.chm262000
f1_keywords:
- vbaof11.chm262000
ms.prod: office
api_name:
- Office.SmartDocument
ms.assetid: b56a86eb-a031-d50b-905e-ef8b91914d61
ms.date: 06/08/2017
---


# SmartDocument Object (Office)

The  **SmartDocument** property of the **Document** object in Microsoft Word and the **Workbook** object in Microsoft Excel returns a **SmartDocument** object.


## Remarks

Use the  **SmartDocument** object to manage the XML expansion pack attached to the active document.

Use the  **SmartDocument** object's **SolutionID** and **SolutionURI** properties to retrieve information about the XML expansion pack attached to the active document or workbook. Use the **PickSolution** method to allow the user to select an available XML expansion pack from a list to attach to the active document or workbook. Use the **RefreshPane** method to refresh the smart document's **Document Actions** task pane.

The  **SmartDocument** object model is available whether or not a document has an XML expansion pack attached. The **SmartDocument** property of the **Document** or **Workbook** objects does not return **Nothing** when the active document has no XML expansion pack attached. Examine the **SolutionID** property to determine whether the active document has an XML expansion pack attached.


## Methods



|**Name**|
|:-----|
|[PickSolution](smartdocument-picksolution-method-office.md)|
|[RefreshPane](smartdocument-refreshpane-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](smartdocument-application-property-office.md)|
|[Creator](smartdocument-creator-property-office.md)|
|[SolutionID](smartdocument-solutionid-property-office.md)|
|[SolutionURL](smartdocument-solutionurl-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
