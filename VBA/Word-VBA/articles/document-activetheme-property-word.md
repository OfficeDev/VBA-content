---
title: Document.ActiveTheme Property (Word)
keywords: vbawd10.chm158007836
f1_keywords:
- vbawd10.chm158007836
ms.prod: word
api_name:
- Word.Document.ActiveTheme
ms.assetid: 2a68899f-8644-c9bb-1d9d-134b132eef91
ms.date: 06/08/2017
---


# Document.ActiveTheme Property (Word)

Returns the name of the active theme plus the theme formatting options for the specified document. Read-only  **String** .


## Syntax

 _expression_ . **ActiveTheme**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **ActiveTheme** property returns "none" if the document doesn't have an active theme. For an explanation of the value returned by this property, see the Name argument of the **[ApplyTheme](document-applytheme-method-word.md)** method. The value returned by this property may not correspond to the theme's display name. To return a theme's display name, use the **[ActiveThemeDisplayName](document-activethemedisplayname-property-word.md)** property.


## Example

This example applies a theme and then displays the name of the active theme plus the theme formatting options for the current document.


```vb
Sub CheckTheme() 
 ActiveDocument.ApplyTheme "artsy 100" 
 MsgBox ActiveDocument.ActiveTheme 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

