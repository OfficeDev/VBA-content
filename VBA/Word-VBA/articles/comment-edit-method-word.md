---
title: Comment.Edit Method (Word)
keywords: vbawd10.chm154993651
f1_keywords:
- vbawd10.chm154993651
ms.prod: word
api_name:
- Word.Comment.Edit
ms.assetid: 94bc4a2e-0b73-af0d-cdac-dec76b1806da
ms.date: 06/08/2017
---


# Comment.Edit Method (Word)

Opens the specified OLE object for editing in the application it was created in.


## Syntax

 _expression_ . **Edit**

 _expression_ Required. A variable that represents a **[Comment](comment-object-word.md)** object.


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


[Comment Object](comment-object-word.md)

