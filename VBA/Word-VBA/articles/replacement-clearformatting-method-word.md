---
title: Replacement.ClearFormatting Method (Word)
keywords: vbawd10.chm162594836
f1_keywords:
- vbawd10.chm162594836
ms.prod: word
api_name:
- Word.Replacement.ClearFormatting
ms.assetid: 3229f741-91f0-1175-5652-96047547d811
ms.date: 06/08/2017
---


# Replacement.ClearFormatting Method (Word)

Removes text and paragraph formatting from the text specified in a replace operation.


## Syntax

 _expression_ . **ClearFormatting**

 _expression_ A variable that represents a **[Replacement](replacement-object-word.md)** object.


## Example

This example clears formatting from the find or replace criteria before replacing the word "Inc." with "incorporated" throughout the active document.


```vb
Sub ClrFmtgReplace() 
 Dim rngTemp As Range 
 Set rngTemp = ActiveDocument.Content 
 With rngTemp.Find 
 .ClearFormatting 
 .Replacement.ClearFormatting 
 .MatchWholeWord = True 
 .Execute FindText:="Inc.", ReplaceWith:="incorporated", _ 
 Replace:=wdReplaceAll 
 End With 
End Sub
```


## See also


#### Concepts


[Replacement Object](replacement-object-word.md)

