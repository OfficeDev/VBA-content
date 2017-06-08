---
title: Options.TypeNReplace Property (Word)
keywords: vbawd10.chm162988457
f1_keywords:
- vbawd10.chm162988457
ms.prod: word
api_name:
- Word.Options.TypeNReplace
ms.assetid: 9696b066-edb5-d7ce-8a4e-ad755acdc738
ms.date: 06/08/2017
---


# Options.TypeNReplace Property (Word)

 **True** for Microsoft Word to replace illegal South Asian characters. Read/write **Boolean** .


## Syntax

 _expression_ . **TypeNReplace**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example instructs Word to replace illegal South Asian characters.


```vb
Sub TypeReplace() 
 Application.Options.TypeNReplace = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

