---
title: Adjustments Object (Word)
ms.prod: word
api_name:
- Word.Adjustments
ms.assetid: ed65525d-2c55-ae2a-ef42-1663b17e5c97
ms.date: 06/08/2017
---


# Adjustments Object (Word)

Contains a collection of adjustment values for the specified AutoShape or WordArt object. Each adjustment value represents one way an adjustment handle can be adjusted. Because some adjustment handles can be adjusted in two ways ? for instance, some handles can be adjusted both horizontally and vertically ? a shape can have more adjustment values than it has adjustment handles. A shape can have up to eight adjustments.


## Remarks

Use the  **Adjustments** property to return an **Adjustments** object. Use **Adjustments** (index), where index is the adjustment value's index number, to return a single adjustment value.

Different shapes have different numbers of adjustment values, different kinds of adjustments change the geometry of a shape in different ways, and different kinds of adjustments have different ranges of valid values.


 **Note**  Because each adjustable shape has a different set of adjustments, the best way to verify the adjustment behavior for a specific shape is to manually create an instance of the shape, make adjustments with the macro recorder turned on, and then examine the recorded code.

The following table summarizes the ranges of valid adjustment values for different types of adjustments. In most cases, if you specify a value that's beyond the range of valid values, the closest valid value will be assigned to the adjustment.



|**Type of Adjustment**|**Valid values**|
|:-----|:-----|
|Linear (horizontal or vertical)|Generally the value 0.0 represents the left or top edge of the shape and the value 1.0 represents the right or bottom edge of the shape. Valid values correspond to valid adjustments you can make to the shape manually. For example, if you can only pull an adjustment handle half way across the shape manually, the maximum value for the corresponding adjustment will be 0.5. For shapes such as callouts, where the values 0.0 and 1.0 represent the limits of the rectangle defined by the starting and ending points of the callout line, negative numbers and numbers greater than 1.0 are valid values.|
|Radial|An adjustment value of 1.0 corresponds to the width of the shape. The maximum value is 0.5, or half way across the shape.|
|Angle|Values are expressed in degrees. If you specify a value outside the range ? 180 to 180, it will be normalized to be within that range.|
The following example adds a right-arrow callout to the active document and sets adjustment values for the callout. Note that although the shape has only three adjustment handles, it has four adjustments. Adjustments three and four both correspond to the handle between the head and neck of the arrow.




```vb
Set rac = ActiveDocument.Shapes _ 
 .AddShape(msoShapeRightArrowCallout, 10, 10, 250, 190) 
With rac.Adjustments 
 .Item(1) = 0.5 'adjusts width of text box 
 .Item(2) = 0.15 'adjusts width of arrow head 
 .Item(3) = 0.8 'adjusts length of arrow head 
 .Item(4) = 0.4 'adjusts width of arrow neck 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

