---
title: Shape.OLEFormat Property (Word)
keywords: vbawd10.chm161481204
f1_keywords:
- vbawd10.chm161481204
ms.prod: word
api_name:
- Word.Shape.OLEFormat
ms.assetid: d558bd26-207c-c308-889e-7316f5027c7f
ms.date: 06/08/2017
---


# Shape.OLEFormat Property (Word)

Returns an  **OLEFormat** object that represents the OLE characteristics (other than linking) for the specified shape, inline shape, or field. Read-only.


## Syntax

 _expression_ . **OLEFormat**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example loops through all the floating shapes on the active document and sets all linked Microsoft Excel worksheets to be updated automatically.


```vb
For Each s In ActiveDocument.Shapes 
 If s.Type = msoLinkedOLEObject Then 
 If s.OLEFormat.ProgID = "Excel.Sheet" Then 
 s.LinkFormat.AutoUpdate = True 
 End If 
 End If 
Next
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

