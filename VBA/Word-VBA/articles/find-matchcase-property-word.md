---
title: Find.MatchCase Property (Word)
keywords: vbawd10.chm162529294
f1_keywords:
- vbawd10.chm162529294
ms.prod: word
api_name:
- Word.Find.MatchCase
ms.assetid: c52c1512-9935-c8a4-4211-5b847771dbe9
ms.date: 06/08/2017
---


# Find.MatchCase Property (Word)

 **True** if the find operation is case sensitive. The default is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **MatchCase**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Remarks

Use the  **[Text](find-text-property-word.md)** property of the **Find** object or use the FindText argument with the **[Execute](find-execute-method-word.md)** method to specify the text to be located in a document.


## Example

This example selects the next occurrence of the word "library" in the selection, regardless of the case.


```vb
With Selection.Find 
 .ClearFormatting 
 .MatchWholeWord = True 
 .MatchCase = False 
 .Execute FindText:="library" 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

