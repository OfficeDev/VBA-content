---
title: Find.MatchWildcards Property (Word)
keywords: vbawd10.chm162529295
f1_keywords:
- vbawd10.chm162529295
ms.prod: word
api_name:
- Word.Find.MatchWildcards
ms.assetid: d2aae410-691e-f718-b888-19e90372d18e
ms.date: 06/08/2017
---


# Find.MatchWildcards Property (Word)

 **True** if the text to find contains wildcards. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchWildcards**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Remarks

The  **MatchWildcards** property corresponds to the **Use wildcards** check box in the **Find and Replace** dialog box ( **Edit** menu).

Use the  **[Text](find-text-property-word.md)** property of the **Find** object or use the FindText argument with the **[Execute](find-execute-method-word.md)** method to specify the text to be located in a document.


## Example

This example finds and selects the next three-letter word that begins with "s" and ends with "t."


```vb
With Selection.Find 
 .ClearFormatting 
 .Text = "s?t" 
 .MatchAllWordForms = False 
 .MatchSoundsLike = False 
 .MatchFuzzy = False 
 .MatchWildcards = True 
 .Execute Format:=False, Forward:=True 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

