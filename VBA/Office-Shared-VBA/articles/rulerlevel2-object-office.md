---
title: RulerLevel2 Object (Office)
ms.prod: office
api_name:
- Office.RulerLevel2
ms.assetid: f1660a26-5990-9524-33f0-a2e3410160f3
ms.date: 06/08/2017
---


# RulerLevel2 Object (Office)

Contains first-line indent and hanging indent information for an outline level.


## Remarks

The  **RulerLevel2** object is a member of the **RulerLevels2** collection. The **RulerLevels2** collection contains a **RulerLevel2** object for each of the five available outline levels.


## Example

Use  `RulerLevels2(index)`, where index is the outline level, to return a single  **RulerLevel2** object. The following example sets the first-line indent and hanging indent for outline level one in body text on the slide master for the active presentation.


```
With ActivePresentation.SlideMaster _ 
 .TextStyles(ppBodyStyle).Ruler2.Levels(1) 
 .FirstMargin = 9 
 .LeftMargin = 54 
End With 

```


## Properties



|**Name**|
|:-----|
|[Application](rulerlevel2-application-property-office.md)|
|[Creator](rulerlevel2-creator-property-office.md)|
|[FirstMargin](rulerlevel2-firstmargin-property-office.md)|
|[LeftMargin](rulerlevel2-leftmargin-property-office.md)|
|[Parent](rulerlevel2-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
