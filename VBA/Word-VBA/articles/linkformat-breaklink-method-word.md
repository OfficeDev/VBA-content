---
title: LinkFormat.BreakLink Method (Word)
keywords: vbawd10.chm154206312
f1_keywords:
- vbawd10.chm154206312
ms.prod: word
api_name:
- Word.LinkFormat.BreakLink
ms.assetid: 19f5f0b5-2536-b6d1-4476-4d46f3d7484e
ms.date: 06/08/2017
---


# LinkFormat.BreakLink Method (Word)

Breaks the link between the source file and the specified OLE object, picture, or linked field.


## Syntax

 _expression_ . **BreakLink**

 _expression_ Required. A variable that represents a **[LinkFormat](linkformat-object-word.md)** object.


## Remarks

After you use this method, the link result won't be automatically updated if the source file is changed.


## Example

This example updates and then breaks the links to any shapes that are linked OLE objects in the active document.


```vb
Dim shapeLoop As Shape 
 
For Each shapeLoop In ActiveDocument.Shapes 
 With shapeLoop 
 If .Type = msoLinkedOLEObject Then 
 .LinkFormat.Update 
 .LinkFormat.BreakLink 
 End If 
 End With 
Next shapeLoop
```


## See also


#### Concepts


[LinkFormat Object](linkformat-object-word.md)

