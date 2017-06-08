---
title: OLEFormat.ProgID Property (PowerPoint)
keywords: vbapp10.chm562005
f1_keywords:
- vbapp10.chm562005
ms.prod: powerpoint
api_name:
- PowerPoint.OLEFormat.ProgID
ms.assetid: 7564f3e1-4e14-9038-a836-5665518b0d09
ms.date: 06/08/2017
---


# OLEFormat.ProgID Property (PowerPoint)

Returns the programmatic identifier (ProgID) for the specified OLE object. Read-only.


## Syntax

 _expression_. **ProgID**

 _expression_ A variable that represents a **OLEFormat** object.


### Return Value

String


## Example

This example loops through all the objects on all the slides in the active presentation and sets all linked Microsoft Office Excel worksheets to be updated manually.


```vb
For Each sld In ActivePresentation.Slides

    For Each sh In sld.Shapes

        If sh.Type = msoLinkedOLEObject Then

            If sh.OLEFormat.ProgID = "Excel.Sheet" Then

                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual

            End If

        End If

    Next

Next
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-powerpoint.md)

