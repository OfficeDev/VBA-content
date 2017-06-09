---
title: Font.Emboss Property (PowerPoint)
keywords: vbapp10.chm575007
f1_keywords:
- vbapp10.chm575007
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Emboss
ms.assetid: 734b5bd7-4b1f-d3b3-d8bd-f73d0bc86f67
ms.date: 06/08/2017
---


# Font.Emboss Property (PowerPoint)

Determines whether the character format is embossed. Read/write.


## Syntax

 _expression_. **Emboss**

 _expression_ A variable that represents an **Font** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Emboss** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The character format is not embossed.|
|**msoTriStateMixed**|The specified text range contains both embossed and unembossed characters.|
|**msoTrue**| The character format is embossed.|

## Example

This example sets the title text on slide one to embossed.


```vb
Application.ActivePresentation.Slides(1).Shapes.Title _
    .TextFrame.TextRange.Font.Emboss = msoTrue
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

