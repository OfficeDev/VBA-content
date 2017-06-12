---
title: Application.SpecialMode Property (Word)
keywords: vbawd10.chm158335006
f1_keywords:
- vbawd10.chm158335006
ms.prod: word
api_name:
- Word.Application.SpecialMode
ms.assetid: aa60d4dc-4abe-e461-12c9-fc8e890536ca
ms.date: 06/08/2017
---


# Application.SpecialMode Property (Word)

 **True** if Microsoft Word is in a special mode (for example, CopyText mode, or MoveText mode). Read-only **Boolean** .


## Syntax

 _expression_ . **SpecialMode**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

Word enters a special copy or move mode if you press F2 or SHIFT+F2 while text is selected.


## Example

This example checks to see whether Word is in a special mode. If it is, ESC is activated before the current selection is deleted and pasted.


```vb
If Application.SpecialMode = True Then SendKeys "ESC" 
With Selection 
 .Cut 
 .EndKey Unit:=wdStory 
 .Paste 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

