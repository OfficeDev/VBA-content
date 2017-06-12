---
title: TextFrame2.HorizontalAnchor Property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.HorizontalAnchor
ms.assetid: 27419e1a-63e6-a08b-2d45-0cd21ada8889
ms.date: 06/08/2017
---


# TextFrame2.HorizontalAnchor Property (Office)

 Returns or sets the horizontal alignment of text in a text frame. Read/write


## Syntax

 _expression_. **HorizontalAnchor**

 _expression_ An expression that returns a **TextFrame2** object.


## Remarks

The value of the  **HorizontalAnchor** property can be one of these **MsoHorizontalAnchor** constants.


||
|:-----|
|**msoAnchorNone**|
|**msoHorizontalAnchorMixed**|
|**msoAnchorCenter**|

## Example

The following code shows how to set the alignment for shape one on slide one to top center.


```
With ActivePresentation.Slides(1).Shapes(1) 
 .TextFrame2.HorizontalAnchor = msoAnchorCenter 
 .TextFrame2.VerticalAnchor = msoAnchorTop 
End With
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-office.md)
#### Other resources


[TextFrame2 Object Members](textframe2-members-office.md)

