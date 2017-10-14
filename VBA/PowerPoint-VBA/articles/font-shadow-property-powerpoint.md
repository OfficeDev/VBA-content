---
title: Font.Shadow Property (PowerPoint)
keywords: vbapp10.chm575006
f1_keywords:
- vbapp10.chm575006
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Shadow
ms.assetid: 37d23e3a-26a7-ba20-1e23-13861090ae79
ms.date: 06/08/2017
---


# Font.Shadow Property (PowerPoint)

Determines whether the specified text has a shadow. Read/write.


## Syntax

 _expression_. **Shadow**

 _expression_ A variable that represents a **Font** object.


## Remarks

The value of the  **Shadow** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified text doesn't have a shadow.|
|**msoTriStateMixed**|Some of the specified text has a shadow and some doesn't.|
|**msoTrue**| The specified text has a shadow.|

## Example

This example adds a shadow to the title text on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes.Title _
    .TextFrame.TextRange.Font.Shadow = True
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

