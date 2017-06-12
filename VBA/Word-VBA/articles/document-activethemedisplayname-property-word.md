---
title: Document.ActiveThemeDisplayName Property (Word)
keywords: vbawd10.chm158007837
f1_keywords:
- vbawd10.chm158007837
ms.prod: word
api_name:
- Word.Document.ActiveThemeDisplayName
ms.assetid: b6689499-80db-12f5-8217-2c982375448b
ms.date: 06/08/2017
---


# Document.ActiveThemeDisplayName Property (Word)

Returns the display name of the active theme for the specified document. Read-only  **String** .


## Syntax

 _expression_ . **ActiveThemeDisplayName**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **ActiveThemeDisplayName** property returns "none" if the document doesn't have an active theme. A theme's display name is the name that appears in the **Theme** dialog box. This name may not correspond to the string you would use to set a default theme or to apply a theme to a document.


## Example

This example returns the display name of the active theme for the current document.


```vb
Sub DisplayThemeName() 
 ActiveDocument.ApplyTheme "artsy 100" 
 MsgBox ActiveDocument.ActiveThemeDisplayName 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

