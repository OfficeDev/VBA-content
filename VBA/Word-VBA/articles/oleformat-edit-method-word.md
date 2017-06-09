---
title: OLEFormat.Edit Method (Word)
keywords: vbawd10.chm154337386
f1_keywords:
- vbawd10.chm154337386
ms.prod: word
api_name:
- Word.OLEFormat.Edit
ms.assetid: 666c20f2-33cf-0655-16f1-914ec0268a1a
ms.date: 06/08/2017
---


# OLEFormat.Edit Method (Word)

Opens the specified OLE object for editing in the application it was created in.


## Syntax

 _expression_ . **Edit**

 _expression_ Required. A variable that represents an **[OLEFormat](oleformat-object-word.md)** object.


## Example

This example opens (for editing) the first embedded OLE object (defined as a shape) on the active document.


```vb
Dim shapesAll As Shapes 
 
Set shapesAll = ActiveDocument.Shapes 
If shapesAll.Count >= 1 Then 
 If shapesAll(1).Type = msoEmbeddedOLEObject Then 
 shapesAll(1).OLEFormat.Edit 
 End If 
End If
```

This example opens (for editing) the first linked OLE object (defined as an inline shape) in the active document.




```vb
Dim colIS As InlineShapes 
 
Set colIS = ActiveDocument.InlineShapes 
If colIS.Count >= 1 Then 
 If colIS(1).Type = wdInlineShapeLinkedOLEObject Then 
 colIS(1).OLEFormat.Edit 
 End If 
End If
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

