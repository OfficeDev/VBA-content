---
title: ValidationIssues.Clear Method (Visio)
keywords: vis_sdr.chm18562410
f1_keywords:
- vis_sdr.chm18562410
ms.prod: visio
api_name:
- Visio.ValidationIssues.Clear
ms.assetid: e3792e98-a47e-2ce2-e1ff-995ccbf645eb
ms.date: 06/08/2017
---


# ValidationIssues.Clear Method (Visio)

Removes all  **[ValidationIssue](validationissue-object-visio.md)** objects from the **[ValidationRules](validationrules-object-visio.md)** collection of the document.


## Syntax

 _expression_ . **Clear**

 _expression_ A variable that represents a **ValidationIssues** object.


### Return Value

 **Nothing**


## Remarks

Calling the  **Clear** method also resets the **[Validation.LastValidatedDate](validation-lastvalidateddate-property-visio.md)** property value to 0 (zero).


## Example

The following sample code is based on code provided by: [David Parker](http://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the  **Clear** method to clear all validation issues from the active document.




```vb

Public Sub Clear_Example()

    ActiveDocument.Validation.Issues.Clear
    
End Sub
```


