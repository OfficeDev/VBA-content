---
title: Find.ClearFormatting Method (Word)
keywords: vbawd10.chm162529311
f1_keywords:
- vbawd10.chm162529311
ms.prod: word
api_name:
- Word.Find.ClearFormatting
ms.assetid: 9b25fb62-13e1-d953-90f2-57059221d820
ms.date: 06/08/2017
---


# Find.ClearFormatting Method (Word)

Removes text and paragraph formatting from the text specified in a find or replace operation.


## Syntax

 _expression_ . **ClearFormatting**

 _expression_ A variable that represents a **[Find](find-object-word.md)** object.


## Example

This example removes formatting from the find criteria before searching through the selection. If the word "Hello" with bold formatting is found, the entire paragraph is selected and copied to the Clipboard.


```vb
Sub ClrFmtgFind() 
 With Selection.Find 
 .ClearFormatting 
 .Font.Bold = True 
 .Execute FindText:="Hello", Format:=True, Forward:=True 
 If .Found = True Then 
 .Parent.Expand Unit:=wdParagraph 
 .Parent.Copy 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Find Object](find-object-word.md)

