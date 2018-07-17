---
title: Presentation.FarEastLineBreakLevel Property (PowerPoint)
keywords: vbapp10.chm583043
f1_keywords:
- vbapp10.chm583043
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.FarEastLineBreakLevel
ms.assetid: fc8354a6-cbd4-d0b4-0b39-a3150afab714
ms.date: 06/08/2017
---


# Presentation.FarEastLineBreakLevel Property (PowerPoint)

Returns or sets the line break based upon Asian character level. Read/write.


## Syntax

 _expression_. **FarEastLineBreakLevel**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

PpFarEastLineBreakLevel


## Remarks

The value of the  **FarEastLineBreakLevel** property can be one of these **PpFarEastLineBreakLevel** constants.


||
|:-----|
|**ppFarEastLineBreakLevelCustom**|
|**ppFarEastLineBreakLevelNormal**|
|**ppFarEastLineBreakLevelStrict**|

## Example

This example sets line break control to use level one kinsoku characters.


```vb
ActivePresentation.FarEastLineBreakLevel = ppFarEastLineBreakLevelNormal
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

