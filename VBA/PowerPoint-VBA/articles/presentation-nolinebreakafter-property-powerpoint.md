---
title: Presentation.NoLineBreakAfter Property (PowerPoint)
keywords: vbapp10.chm583045
f1_keywords:
- vbapp10.chm583045
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.NoLineBreakAfter
ms.assetid: bc9c7fd9-4aa6-b350-4c30-586a237d904a
ms.date: 06/08/2017
---


# Presentation.NoLineBreakAfter Property (PowerPoint)

Returns or sets the characters that cannot end a line. Read/write.


## Syntax

 _expression_. **NoLineBreakAfter**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

String


## Example

This example sets "$", "(", "[", "\", and "{" as characters that cannot end a line.


```vb
With ActivePresentation

    .FarEastLineBreakLevel = ppFarEastLineBreakLevelCustom

    .NoLineBreakAfter =  "$([\{"

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

