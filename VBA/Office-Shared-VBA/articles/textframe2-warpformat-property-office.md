---
title: TextFrame2.WarpFormat Property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.WarpFormat
ms.assetid: 83993a3d-a594-e3bc-47ca-47f50be143b7
ms.date: 06/08/2017
---


# TextFrame2.WarpFormat Property (Office)

Returns or sets the warp format (how the text is warped) for the specified text frame. Read/write


## Syntax

 _expression_. **WarpFormat**

 _expression_ An expression that returns a **TextFrame2** object.


## Remarks

The value of the WarpFormat property can be one of the MsoWarpFormat constants.


## Example

The following code shows how to set the warp format for shape one on slide one of the active presentation.


```
Public Sub WarpFormat_Example() 
 
 Dim pptSlide As Slide 
 Set pptSlide = ActivePresentation.Slides(1) 
 pptSlide.Shapes(1).TextFrame2.WarpFormat = msoWarpFormat15 
 
End Sub 

```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-office.md)
#### Other resources


[TextFrame2 Object Members](textframe2-members-office.md)

