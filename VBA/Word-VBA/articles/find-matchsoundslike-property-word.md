---
title: Find.MatchSoundsLike Property (Word)
keywords: vbawd10.chm162529296
f1_keywords:
- vbawd10.chm162529296
ms.prod: word
api_name:
- Word.Find.MatchSoundsLike
ms.assetid: 81c341a7-40a8-7022-78d5-a8ed8ad407b1
ms.date: 06/08/2017
---


# Find.MatchSoundsLike Property (Word)

 **True** if words that sound similar to the text to find are returned by the find operation. Read/write **Boolean** .


## Syntax

 _expression_ . **MatchSoundsLike**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Remarks

Use the  **[Text](find-text-property-word.md)** property of the **Find** object or use the FindText argument with the **[Execute](find-execute-method-word.md)** method to specify the text to be located in a document.


## Example

This example selects the next word that sounds like the word "fun" (for instance, "funny") in the selection.


```vb
With Selection.Find 
 .ClearFormatting 
 .Text = "fun" 
 .MatchFuzzy = False 
 .MatchSoundsLike = True 
 .Execute Format:=False, Forward:=True, Wrap:=wdFindContinue 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

