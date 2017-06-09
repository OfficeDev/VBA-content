---
title: OLEFormat.IconLabel Property (Word)
keywords: vbawd10.chm154337290
f1_keywords:
- vbawd10.chm154337290
ms.prod: word
api_name:
- Word.OLEFormat.IconLabel
ms.assetid: 8cf2aaf3-0ce0-80b4-a5ad-2561f1af4457
ms.date: 06/08/2017
---


# OLEFormat.IconLabel Property (Word)

Returns or sets the text displayed below the icon for an OLE object. Read/write  **String** .


## Syntax

 _expression_ . **IconLabel**

 _expression_ An expression that returns an **[OLEFormat](oleformat-object-word.md)** object.


## Example

This example changes the text below the icon for the first shape in the selection.


```vb
Dim olefTemp As OLEFormat 
 
If Selection.ShapeRange.Count >= 1 Then 
 Set olefTemp = Selection.ShapeRange(1).OLEFormat 
 With olefTemp 
 .DisplayAsIcon = True 
 .IconLabel = "My Icon" 
 End With 
End If
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

