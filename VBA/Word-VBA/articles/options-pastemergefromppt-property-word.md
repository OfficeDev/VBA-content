---
title: Options.PasteMergeFromPPT Property (Word)
keywords: vbawd10.chm162988465
f1_keywords:
- vbawd10.chm162988465
ms.prod: word
api_name:
- Word.Options.PasteMergeFromPPT
ms.assetid: 5e0b04ba-5dce-a3cf-9bc8-672f55b5b10e
ms.date: 06/08/2017
---


# Options.PasteMergeFromPPT Property (Word)

 **True** to merge text formatting when pasting from Microsoft PowerPoint. Read/write **Boolean** .


## Syntax

 _expression_ . **PasteMergeFromPPT**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically merge text formatting when pasting content from PowerPoint if the option has been disabled.


```vb
Sub AdjustPPTFormatting() 
 With Options 
 If .PasteMergeFromPPT = False Then 
 .PasteMergeFromPPT = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

