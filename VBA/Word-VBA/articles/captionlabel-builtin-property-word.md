---
title: CaptionLabel.BuiltIn Property (Word)
keywords: vbawd10.chm158924801
f1_keywords:
- vbawd10.chm158924801
ms.prod: word
api_name:
- Word.CaptionLabel.BuiltIn
ms.assetid: 1df0a271-2792-0813-f45d-2b076afa0a3a
ms.date: 06/08/2017
---


# CaptionLabel.BuiltIn Property (Word)

 **True** if the specified caption label is one of the built-in caption labels in Word. Read-only **Boolean** .


## Syntax

 _expression_ . **BuiltIn**

 _expression_ A variable that represents a **[CaptionLabel](captionlabel-object-word.md)** object.


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

This example checks all the caption labels that have been created in the application. When it finds a caption label that isn't built in, it displays the name of the label.




```vb
Dim clLoop As CaptionLabel 
 
For Each clLoop in CaptionLabels 
 If clLoop.BuiltIn = False Then 
 Msgbox clLoop.Name 
 End If 
Next clLoop
```


## See also


#### Concepts


[CaptionLabel Object](captionlabel-object-word.md)

