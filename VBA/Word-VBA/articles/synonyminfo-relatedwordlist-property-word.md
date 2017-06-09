---
title: SynonymInfo.RelatedWordList Property (Word)
keywords: vbawd10.chm161153034
f1_keywords:
- vbawd10.chm161153034
ms.prod: word
api_name:
- Word.SynonymInfo.RelatedWordList
ms.assetid: 7126c71c-6308-9b4b-89c7-6762e01fc591
ms.date: 06/08/2017
---


# SynonymInfo.RelatedWordList Property (Word)

Returns a list of words related to the specified word or phrase. The list is returned as an array of strings. Read-only  **Variant** .


## Syntax

 _expression_ . **RelatedWordList**

 _expression_ An expression that returns a **[SynonymInfo](synonyminfo-object-word.md)** object.


## Example

This example checks to see whether any related words were found for the third word in the active document. If so, the meanings are displayed in a series of message boxes. If there are no related words found, this is stated in a message box.


```vb
Set synInfo = ActiveDocument.Words(3).SynonymInfo 
If synInfo.Found = True Then 
 relList = synInfo.RelatedWordList 
 If UBound(relList) <> 0 Then 
 For intCount = 1 To UBound(relList) 
 Msgbox relList(intCount) 
 Next intCount 
 Else 
 Msgbox "There were no related words found." 
 End If 
End If
```


## See also


#### Concepts


[SynonymInfo Object](synonyminfo-object-word.md)

