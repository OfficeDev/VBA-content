---
title: RulerLevel.FirstMargin Property (PowerPoint)
keywords: vbapp10.chm572003
f1_keywords:
- vbapp10.chm572003
ms.prod: powerpoint
api_name:
- PowerPoint.RulerLevel.FirstMargin
ms.assetid: 9149006e-d518-bb68-4244-b6047387b672
ms.date: 06/08/2017
---


# RulerLevel.FirstMargin Property (PowerPoint)

Returns or sets the first-line indent for the specified outline level, in points. Read/write.


## Syntax

 _expression_. **FirstMargin**

 _expression_ A variable that represents a **RulerLevel** object.


### Return Value

Single


## Remarks

If a paragraph begins with a bullet, the bullet position is determined by the  **FirstMargin** property, and the position of the first text character in the paragraph is determined by the **LeftMargin** property.


 **Note**  The  **[RulerLevels](rulerlevels-object-powerpoint.md)** collection contains five **RulerLevel** objects, each of which corresponds to one of the possible outline levels. The **FirstMargin** property value for the **RulerLevel** object that corresponds to the first outline level has a valid range of (-9.0 to 4095.875). The valid range for the **FirstMargin** property values for the **RulerLevel** objects that correspond to the second through the fifth outline levels are determined as follows:


- The maximum value is always 4095.875.
    
- The minimum value is the maximum assigned value between the  **FirstMargin** property and **LeftMargin** property of the previous level plus 9.
    
You can use the following equations to determine the minimum value for the  **FirstMargin** property. Index, the index number of the **RulerLevel** object, indicates the object's corresponding outline level. To determine the minimum **FirstMargin** property values for the **RulerLevel** objects that correspond to the second through the fifth outline levels, substitute 2, 3, 4, or 5 for the index placeholder.

Minimum(RulerLevel(index). **FirstMargin** ) = Maximum(RulerLevel(index -1). **FirstMargin**, RulerLevel(index -1).LeftMargin) + 9

Minimum (RulerLevel(index).LeftMargin) = Maximum(RulerLevel(index -1). **FirstMargin**, RulerLevel(index -1).LeftMargin) + 9


## Example

This example sets the first-line indent and hanging indent for outline level one in body text on the slide master for the active presentation.


```vb
With Application.ActivePresentation _
        .SlideMaster.TextStyles(ppBodyStyle)
    With .Ruler.Levels(1)
        .FirstMargin = 9
        .LeftMargin = 54
    End With
End With
```


## See also


#### Concepts


[RulerLevel Object](rulerlevel-object-powerpoint.md)

