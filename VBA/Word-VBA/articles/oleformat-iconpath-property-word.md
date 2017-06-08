---
title: OLEFormat.IconPath Property (Word)
keywords: vbawd10.chm154337288
f1_keywords:
- vbawd10.chm154337288
ms.prod: word
api_name:
- Word.OLEFormat.IconPath
ms.assetid: 787bfe10-943c-e470-23e3-10abec89e606
ms.date: 06/08/2017
---


# OLEFormat.IconPath Property (Word)

Returns the path of the file in which the icon for an OLE object is stored. Read-only  **String** .


## Syntax

 _expression_ . **IconPath**

 _expression_ An expression that returns an **[OLEFormat](oleformat-object-word.md)** object.


## Example

This example displays the path for each embedded OLE object that's displayed as an icon on the active document.


```vb
Dim shapeLoop As Shape 
 
For Each shapeLoop In ActiveDocument.Shapes 
 If shapeLoop.Type = msoEmbeddedOLEObject Then 
 If shapeLoop.OLEFormat.DisplayAsIcon = True Then _ 
 Msgbox shapeLoop.OLEFormat.IconPath 
 End If 
Next shapeLoop
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

