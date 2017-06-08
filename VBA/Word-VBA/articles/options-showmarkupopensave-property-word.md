---
title: Options.ShowMarkupOpenSave Property (Word)
keywords: vbawd10.chm162988490
f1_keywords:
- vbawd10.chm162988490
ms.prod: word
api_name:
- Word.Options.ShowMarkupOpenSave
ms.assetid: 7e622cce-2465-48fd-08c0-3385cbc36d55
ms.date: 06/08/2017
---


# Options.ShowMarkupOpenSave Property (Word)

Returns or sets a  **Boolean** that represents whether Microsoft Word displays hidden markup when opening or saving a file.


## Syntax

 _expression_ . **ShowMarkupOpenSave**

 _expression_ An expression that returns an **Options** object.


## Remarks

The  **ShowMarkupOpenSave** property corresponds to the **Make hidden markup visible when opening or saving** option in the **Security** tab of the **Options** dialog box.


## Example

The following example enables the Make hidden markup visible when opening or saving option.


```vb
Options.ShowMarkupOpenSave = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)

