---
title: ColorFormat.SchemeColor Property (Publisher)
keywords: vbapb10.chm2555910
f1_keywords:
- vbapb10.chm2555910
ms.prod: publisher
api_name:
- Publisher.ColorFormat.SchemeColor
ms.assetid: 8b02c85c-a976-7b10-c4ea-6f881d702b55
ms.date: 06/08/2017
---


# ColorFormat.SchemeColor Property (Publisher)

Specifies the color of the current color scheme. Read/write.


## Syntax

 _expression_. **SchemeColor**

 _expression_A variable that represents a  **ColorFormat** object.


### Return Value

PbSchemeColorIndex


## Remarks

The  **SchemeColor** property value can be one of the **[PbSchemeColorIndex](pbschemecolorindex-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

The following example sets the color of the text in shape one on page one of the active publication to accent color five in the current color scheme.


```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Font.Color.SchemeColor =
```


```
pbSchemeColorAccent5
```


