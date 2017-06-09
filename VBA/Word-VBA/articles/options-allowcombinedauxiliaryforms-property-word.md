---
title: Options.AllowCombinedAuxiliaryForms Property (Word)
keywords: vbawd10.chm162988371
f1_keywords:
- vbawd10.chm162988371
ms.prod: word
api_name:
- Word.Options.AllowCombinedAuxiliaryForms
ms.assetid: c692e1de-7b89-7255-7fba-6c6bdd472a0a
ms.date: 06/08/2017
---


# Options.AllowCombinedAuxiliaryForms Property (Word)

 **True** if Microsoft Word ignores auxiliary verb forms when checking spelling in a Korean language document. Read/write **Boolean** .


## Syntax

 _expression_ . **AllowCombinedAuxiliaryForms**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

For more information on using Word with Asian languages, see Word features for Asian languages .


## Example

This example asks the user whether Microsoft Word should ignore auxiliary verb forms when checking spelling in a Korean language document.


```vb
If Options.AllowCombinedAuxiliaryForms = False Then 
 x = MsgBox("Do you want to ignore auxiliary " _ 
 &; "verb forms when checking spelling?", _ 
 vbYesNo) 
 If x = vbYes Then 
 Options.AllowCombinedAuxiliaryForms = True 
 MsgBox "Auxiliary verb forms will be ignored!" 
 End If 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

