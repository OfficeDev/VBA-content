---
title: LinkFormat.Update Method (PowerPoint)
keywords: vbapp10.chm563005
f1_keywords:
- vbapp10.chm563005
ms.prod: powerpoint
api_name:
- PowerPoint.LinkFormat.Update
ms.assetid: c1ce2e2f-53ca-9c64-4ce5-1e0d0bed6c54
ms.date: 06/08/2017
---


# LinkFormat.Update Method (PowerPoint)

Updates the specified linked OLE object. 


## Syntax

 _expression_. **Update**

 _expression_ A variable that represents an **LinkFormat** object.


## Remarks

To update all the links in a presentation at once, use the  **[UpdateLinks](presentation-updatelinks-method-powerpoint.md)** method.


## Example

This example updates all linked OLE objects in the active presentation.


```vb
For Each sld In ActivePresentation.Slides

    For Each sh In sld.Shapes

        If sh.Type = msoLinkedOLEObject Then

            sh.LinkFormat.Update

        End If

    Next

Next
```


## See also


#### Concepts


[LinkFormat Object](linkformat-object-powerpoint.md)

