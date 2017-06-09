---
title: ShapeRange.OLEFormat Property (PowerPoint)
keywords: vbapp10.chm548044
f1_keywords:
- vbapp10.chm548044
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.OLEFormat
ms.assetid: ff454e81-5c55-5deb-9816-0eb06b236a95
ms.date: 06/08/2017
---


# ShapeRange.OLEFormat Property (PowerPoint)

Returns an  **[OLEFormat](oleformat-object-powerpoint.md)** object that contains OLE formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent OLE objects. Read-only.


## Syntax

 _expression_. **OLEFormat**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

OLEFormat


## Example

This example loops through all the objects on all the slides in the active presentation and sets all linked Microsoft Word documents to be updated manually.


```vb
For Each sld In ActivePresentation.Slides

    For Each sh In sld.Shapes

        If sh.Type = msoLinkedOLEObject Then

            If sh.OLEFormat.ProgID = "Word.Document" Then

                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual

            End If

        End If

    Next

Next
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

