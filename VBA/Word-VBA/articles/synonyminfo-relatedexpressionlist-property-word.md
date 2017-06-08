---
title: SynonymInfo.RelatedExpressionList Property (Word)
keywords: vbawd10.chm161153033
f1_keywords:
- vbawd10.chm161153033
ms.prod: word
api_name:
- Word.SynonymInfo.RelatedExpressionList
ms.assetid: a7ce0fa7-cffb-a569-0a2a-894ede869f26
ms.date: 06/08/2017
---


# SynonymInfo.RelatedExpressionList Property (Word)

Returns a list of expressions related to the specified word or phrase. The list is returned as an array of strings. Read-only  **Variant** .


## Syntax

 _expression_ . **RelatedExpressionList**

 _expression_ An expression that returns a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Remarks

Typically, there are very few related expressions found in the thesaurus.


## Example

This example checks to see whether any related expressions were found for the selection. If so, the meanings are displayed in a series of message boxes. If none were found, this is stated in a message box.


```vb
Set synInfo = Selection.Range.SynonymInfo 
If synInfo.Found = True Then 
 relList = synInfo.RelatedExpressionList 
 If UBound(relList) <> 0 Then 
 For intCount = 1 To UBound(relList) 
 Msgbox relList(intCount) 
 Next intCount 
 Else 
 Msgbox "There were no related expressions found." 
 End If 
End If
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)

