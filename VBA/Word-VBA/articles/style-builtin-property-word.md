---
title: Style.BuiltIn Property (Word)
keywords: vbawd10.chm153878532
f1_keywords:
- vbawd10.chm153878532
ms.prod: word
api_name:
- Word.Style.BuiltIn
ms.assetid: dee6db94-7f87-3cfc-de76-b6bda8911cce
ms.date: 06/08/2017
---


# Style.BuiltIn Property (Word)

 **True** if the specified object is one of the built-in styles or caption labels in Word. Read-only **Boolean** .


## Syntax

 _expression_ . **BuiltIn**

 _expression_ A variable that represents a **[Style](style-object-word.md)** object.


## Remarks

You can specify built-in styles across all languages by using the  **WdBuiltinStyle** constants or within a language by using the style name for the language version of Word. For example, if you specify U.S. English in your Microsoft Office language settings, the following statements are equivalent:


```vb
ActiveDocument.Styles(wdStyleHeading1) 
ActiveDocument.Styles("Heading 1")
```


## Example

This example checks all the styles in the active document. When it finds a style that isn't built in, it displays the name of the style.


```vb
Dim styleLoop As Style 
 
For Each styleLoop in ActiveDocument.Styles 
 If styleLoop.BuiltIn = False Then 
 Msgbox styleLoop.NameLocal 
 End If 
Next styleLoop
```


## See also


#### Concepts


[Style Object](style-object-word.md)

