---
title: TextFrame2.Orientation Property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.Orientation
ms.assetid: 529b71d3-d653-61c6-eb0a-69b2f3910d0a
ms.date: 06/08/2017
---


# TextFrame2.Orientation Property (Office)

Returns or sets text orientation. Read/write


## Syntax

 _expression_. **Orientation**

 _expression_ An expression that returns a **TextFrame2** object.


## Remarks

The value of the Orientation property can be one of these MsoTextOrientation constants.


-  **msoTextOrientationDownward**
    
-  **msoTextOrientationHorizontal**
    
-  **msoTextOrientationHorizontalRotatedFarEast**
    
-  **msoTextOrientationMixed**
    
-  **msoTextOrientationUpward**
    
-  **msoTextOrientationVertical**
    
-  **msoTextOrientationVerticalFarEast**
    

## Example

This example shows how to orient the text horizontally in shape one on slide one in the active presentation. 


```
Dim pptSlide As Slide 
Set pptSlide = ActivePresentation.Slides(1) 
pptSlide.Shapes(1).TextFrame2.Orientation = msoTextOrientationHorizontal
```


 **Note**  Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-office.md)
#### Other resources


[TextFrame2 Object Members](textframe2-members-office.md)

