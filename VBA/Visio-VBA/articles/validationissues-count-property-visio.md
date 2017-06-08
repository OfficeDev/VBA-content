---
title: ValidationIssues.Count Property (Visio)
keywords: vis_sdr.chm18513330
f1_keywords:
- vis_sdr.chm18513330
ms.prod: visio
api_name:
- Visio.ValidationIssues.Count
ms.assetid: 7077d75d-640c-32ee-fdf3-1be37407ab94
ms.date: 06/08/2017
---


# ValidationIssues.Count Property (Visio)

Returns the number of  **[ValidationIssue](validationissue-object-visio.md)** objects in the collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **[ValidationIssues](validationissues-object-visio.md)** object.


### Return Value

 **Long**


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Count** method to determine how many validation issues exist in the collection of validation issues in the active document.


```vb
Set vsoDocument = Visio.ActiveDocument 
Set vsoIssues = vsoDocument.Validation.Issues
intIssueTotal = vsoIssues.Count
```


