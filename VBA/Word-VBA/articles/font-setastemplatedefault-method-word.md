---
title: Font.SetAsTemplateDefault Method (Word)
keywords: vbawd10.chm156368999
f1_keywords:
- vbawd10.chm156368999
ms.prod: word
api_name:
- Word.Font.SetAsTemplateDefault
ms.assetid: 91c32f0a-52bd-cddf-9ce1-362bc205d234
ms.date: 06/08/2017
---


# Font.SetAsTemplateDefault Method (Word)

Sets the specified font formatting as the default for the active document and all new documents based on the active template.


## Syntax

 _expression_ . **SetAsTemplateDefault**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

The default font formatting is stored in the Normal style.


## Example

This example sets the character formatting in the selection as the default.


```
Selection.Font.SetAsTemplateDefault
```


## See also


#### Concepts


[Font Object](font-object-word.md)

