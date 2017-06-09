---
title: Master.PatternFlags Property (Visio)
keywords: vis_sdr.chm10714065
f1_keywords:
- vis_sdr.chm10714065
ms.prod: visio
api_name:
- Visio.Master.PatternFlags
ms.assetid: cf7d5e0e-802e-c65b-6260-eaf68dfe6eb4
ms.date: 06/08/2017
---


# Master.PatternFlags Property (Visio)

Determines whether a master behaves as a custom pattern. Read/write.


## Syntax

 _expression_ . **PatternFlags**

 _expression_ A variable that represents a **Master** object.


### Return Value

Integer


## Remarks

Microsoft Visio allows a master to be used as a custom line pattern, line end, or fill pattern.

The  **PatternFlags** property determines whether you can use a master as a pattern (non-zero); whether it is a line, fill, or line end pattern; and which pattern mode to use when applying it to shapes.

If you can use the  **PatternFlags** property as a pattern (non-zero), the property can include a combination of the following bits.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visMasIsLinePat**|&;H1 |Line pattern |
| **visMasIsLineEnd**|&;H2 |Line end pattern |
| **visMasIsFillPat**|&;H4 |Fill pattern |
If  **visMasIsLinePat** is selected, the pattern mode should be one of the following values.



|**Constant **|**Value **|
|:-----|:-----|
| **visMasLPTileDeform**|&;H0 |
| **visMasLPTile**|&;H10 |
| **visMasLPStretch**|&;H20 |
| **visMasLPAnnotate**|&;H30 |
In addition,  **visMasLPScale** (&;H40) can optionally be included in the **PatternFlags** property value.

If  **visMasIsLineEnd** is selected, the pattern mode should be one of the following values.



|**Constant **|**Value **|
|:-----|:-----|
| **visMasLEDefault**|&;H0 |
| **visMasLEUpright**|&;H100 |
In addition,  **visMasLEScale** (&;H400) can optionally be included in the **PatternFlags** property value.

If  **visMasIsFillPat** is selected, the pattern mode should be one of the following values.



|**Constant **|**Value **|
|:-----|:-----|
| **visMasFPTile**|&;H0 |
| **visMasFPCenter**|&;H1000 |
| **visMasFPStretch**|&;H2000 |
In addition,  **visMasFPScale** (&;H4000) can optionally be included in the **PatternFlags** property value.


