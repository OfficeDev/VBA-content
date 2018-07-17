---
title: Options.PrintBackgrounds Property (Word)
keywords: vbawd10.chm162988488
f1_keywords:
- vbawd10.chm162988488
ms.prod: word
api_name:
- Word.Options.PrintBackgrounds
ms.assetid: 81c15f4a-c6ea-9be2-8f3e-bb215ee7af4e
ms.date: 06/08/2017
---


# Options.PrintBackgrounds Property (Word)

Returns a  **Boolean** that represents whether background colors and images are printed when a document is printed.


## Syntax

 _expression_ . **PrintBackgrounds**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

 **True** indicates that background colors and images are printed. **False** indicates that background colors and images are not printed.


## Example

The following example specifies that when documents are printed background colors and images will also be printed.


```vb
Options.PrintBackgrounds = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)

