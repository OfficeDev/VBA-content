---
title: Style.NoProofing Property (Word)
keywords: vbawd10.chm153878546
f1_keywords:
- vbawd10.chm153878546
ms.prod: word
api_name:
- Word.Style.NoProofing
ms.assetid: dbfc95ea-160a-bda9-e7e8-b73ae2314228
ms.date: 06/08/2017
---


# Style.NoProofing Property (Word)

 **True** if the spelling and grammar checker ignores text formatted with this style. Read/write **Long** .


## Syntax

 _expression_ . **NoProofing**

 _expression_ A variable that represents a **[Style](style-object-word.md)** object.


## Example

This example sets the spelling and grammar checker to ignore any text in the active document formatted with the style named "Test".


```vb
ActiveDocument.Styles("Test").NoProofing = True
```


## See also


#### Concepts


[Style Object](style-object-word.md)

