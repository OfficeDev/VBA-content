---
title: LinkFormat.SourcePath Property (Word)
keywords: vbawd10.chm154206213
f1_keywords:
- vbawd10.chm154206213
ms.prod: word
api_name:
- Word.LinkFormat.SourcePath
ms.assetid: c5aa7b91-7c65-b9d7-3e5e-8eb203340d08
ms.date: 06/08/2017
---


# LinkFormat.SourcePath Property (Word)

Returns the path of the source file for the specified linked OLE object, picture, or field. Read-only  **String** .


## Syntax

 _expression_ . **SourcePath**

 _expression_ An expression that returns a **[LinkFormat](linkformat-object-word.md)** object.


## Remarks

The path doesn't include a trailing character (for example, "C:\MSOffice"). Use the  **[PathSeparator](application-pathseparator-property-word.md)** property to add the character that separates folders and drive letters. Use the **[SourceName](linkformat-sourcename-property-word.md)** property to return the file name without the path and use the **[SourceFullName](linkformat-sourcefullname-property-word.md)** property to return the path and file name together.


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

