---
title: LinkFormat.AutoUpdate Property (Word)
keywords: vbawd10.chm154206209
f1_keywords:
- vbawd10.chm154206209
ms.prod: word
api_name:
- Word.LinkFormat.AutoUpdate
ms.assetid: 39525118-e17e-d19e-33b8-98dc52d895f2
ms.date: 06/08/2017
---


# LinkFormat.AutoUpdate Property (Word)

 **True** if the specified link is updated automatically when the container file is opened or when the source file is changed. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoUpdate**

 _expression_ A variable that represents a **[LinkFormat](linkformat-object-word.md)** object.


## Example

This example updates any shapes in the active document that are linked OLE objects if Word isn't set to update links automatically.


```vb
Dim shapeLoop as Shape 
 
For Each shapeLoop In ActiveDocument.Shapes 
 With shapeLoop 
 If .Type = msoLinkedOLEObject Then 
 If .LinkFormat.AutoUpdate = False Then 
 .LinkFormat.Update 
 End If 
 End If 
 End With 
Next s
```

This example updates any fields in the active document that aren't updated automatically.




```vb
Dim fieldLoop as Field 
 
For Each fieldLoop In ActiveDocument.Fields 
 If fieldLoop.LinkFormat.AutoUpdate = False Then _ 
 fieldLoop.LinkFormat.Update 
Next fieldLoop
```


## See also


#### Concepts


[LinkFormat Object](linkformat-object-word.md)

