---
title: Options.AllowCompoundNounProcessing Property (Word)
keywords: vbawd10.chm162988379
f1_keywords:
- vbawd10.chm162988379
ms.prod: word
api_name:
- Word.Options.AllowCompoundNounProcessing
ms.assetid: 78da1977-2d44-7686-5e31-2e7c340f726f
ms.date: 06/08/2017
---


# Options.AllowCompoundNounProcessing Property (Word)

 **True** if Microsoft Word ignores compound nouns when checking spelling in a Korean language document. Read/write **Boolean** .


## Syntax

 _expression_ . **AllowCompoundNounProcessing**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

For more information on using Word with Asian languages, see Word features for Asian languages .


## Example

This example asks the user whether Microsoft Word should ignore compound nouns when checking spelling in a Korean language document.


```vb
If Options.AllowCompoundNounProcessing = False Then 
 x = MsgBox("Do you want to ignore compound " _ 
 &; "nouns when checking spelling?", _ 
 vbYesNo) 
 If x = vbYes Then 
 Options.AllowCompoundNounProcessing = True 
 MsgBox "Compound nouns will be ignored!" 
 End If 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

