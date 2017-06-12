---
title: Validation.Issues Property (Visio)
keywords: vis_sdr.chm18062720
f1_keywords:
- vis_sdr.chm18062720
ms.prod: visio
api_name:
- Visio.Validation.Issues
ms.assetid: a6d79208-9e94-733a-8432-1cd9784e8dc2
ms.date: 06/08/2017
---


# Validation.Issues Property (Visio)

Returns the collection of all the validation issues in the document. Read-only.


## Syntax

 _expression_ . **Issues**

 _expression_ A variable that represents a **[Validation](validation-object-visio.md)** object.


### Return Value

 **[ValidationIssues](validationissues-object-visio.md)**


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Issues** property to get an object that represents the collection of all the validation issues in the document.


```vb
Dim vsoDocument As Visio.Document
Dim vsoIssues As Visio.ValidationIssues

Set vsoDocument = Visio.ActiveDocument 
Set vsoIssues = vsoDocument.Validation.Issues
```


