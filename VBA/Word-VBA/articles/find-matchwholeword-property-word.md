---
title: Find.MatchWholeWord Property (Word)
keywords: vbawd10.chm162529297
f1_keywords:
- vbawd10.chm162529297
ms.prod: word
api_name:
- Word.Find.MatchWholeWord
ms.assetid: a4ce7e5f-c84b-b13a-e21c-14051a0f4a6a
ms.date: 06/08/2017
---


# Find.MatchWholeWord Property (Word)

 **True** if the find operation locates only entire words and not text that's part of a larger word. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchWholeWord**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Remarks

Use the  **[Text](find-text-property-word.md)** property of the **Find** object or the FindText argument with the **[Execute](find-execute-method-word.md)** method to specify the text to be located in a document.


## Example

This example clears all formatting from the find and replace criteria before replacing the word "Inc." with "incorporated" throughout the active document.


```vb
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Replacement.ClearFormatting 
 .MatchWholeWord = True 
 .Execute FindText:="Inc.", _ 
 ReplaceWith:="incorporated", Replace:=wdReplaceAll 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

