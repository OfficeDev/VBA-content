---
title: Shape.OLEFormat Property (PowerPoint)
keywords: vbapp10.chm547044
f1_keywords:
- vbapp10.chm547044
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.OLEFormat
ms.assetid: d9353732-0b91-ae53-a468-07a57359295d
ms.date: 06/08/2017
---


# Shape.OLEFormat Property (PowerPoint)

Returns an  **[OLEFormat](oleformat-object-powerpoint.md)** object that contains OLE formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent OLE objects. Read-only.


## Syntax

 _expression_. **OLEFormat**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

