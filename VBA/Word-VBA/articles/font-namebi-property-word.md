---
title: Font.NameBi Property (Word)
keywords: vbawd10.chm156369059
f1_keywords:
- vbawd10.chm156369059
ms.prod: word
api_name:
- Word.Font.NameBi
ms.assetid: 436dd5c5-a79d-265e-9929-f30c5a05e85e
ms.date: 06/08/2017
---


# Font.NameBi Property (Word)

Returns or sets the name of the font in a right-to-left language document. Read/write  **String** .


## Syntax

 _expression_ . **NameBi**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Example

This example formats the selection with Arial font.


```vb
With Selection.Font 
 .NameBi = "Arial" 
End With
```


## See also


#### Concepts


[Font Object](font-object-word.md)

