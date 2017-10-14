---
title: LinkFormat.SourceName Property (Word)
keywords: vbawd10.chm154206212
f1_keywords:
- vbawd10.chm154206212
ms.prod: word
api_name:
- Word.LinkFormat.SourceName
ms.assetid: 1befe8a0-29f4-21cc-e2cb-03ce018db620
ms.date: 06/08/2017
---


# LinkFormat.SourceName Property (Word)

Returns the name of the source file for the specified linked OLE object, picture, or field. Read-only  **String** .


## Syntax

 _expression_ . **SourceName**

 _expression_ An expression that returns a **[LinkFormat](linkformat-object-word.md)** object.


## Remarks

This property doesn't return the path for the source file.


## Example

This example returns the path and name of the source file for any shapes on the active document that are linked OLE objects.


```vb
For Each s In ActiveDocument.Shapes 
 If s.Type = msoLinkedOLEObject Then 
 Msgbox s.LinkFormat.SourcePath &; "\" _ 
 &; s.LinkFormat.SourceName 
 End If 
Next s
```


## See also


#### Concepts


[LinkFormat Object](linkformat-object-word.md)

