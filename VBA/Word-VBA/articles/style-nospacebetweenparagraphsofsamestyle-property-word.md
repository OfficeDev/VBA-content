---
title: Style.NoSpaceBetweenParagraphsOfSameStyle Property (Word)
keywords: vbawd10.chm153878548
f1_keywords:
- vbawd10.chm153878548
ms.prod: word
api_name:
- Word.Style.NoSpaceBetweenParagraphsOfSameStyle
ms.assetid: 922aa621-0c52-cc7e-9713-1e129bba77c0
ms.date: 06/08/2017
---


# Style.NoSpaceBetweenParagraphsOfSameStyle Property (Word)

 **True** for Microsoft Word to remove spacing between paragraphs that are formatted using the same style. Read/write **Boolean** .


## Syntax

 _expression_ . **NoSpaceBetweenParagraphsOfSameStyle**

 _expression_ An expression that returns a **[Style](style-object-word.md)** object.


## Example


```vb
Sub NoSpace() 
 ActiveDocument.Styles("List 1") _ 
 .NoSpaceBetweenParagraphsOfSameStyle = True 
End Sub
```


## See also


#### Concepts


[Style Object](style-object-word.md)

