---
title: Font.Bold Property (PowerPoint)
keywords: vbapp10.chm575004
f1_keywords:
- vbapp10.chm575004
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Bold
ms.assetid: 13e81c46-5ae7-21ee-58e1-5ab23de552d5
ms.date: 06/08/2017
---


# Font.Bold Property (PowerPoint)

Determines whether the character format is bold. Read/write.


## Syntax

 _expression_. **Bold**

 _expression_ A variable that represents a **Font** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Bold** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The character format is not bold.|
|**msoTriStateMixed**|The specified text range contains both bold and nonbold characters.|
|**msoTrue**| The character format is bold.|

## Example

This example sets characters one through five in the title on slide one to bold.


```vb
Set myT = Application.ActivePresentation.Slides(1).Shapes.Title

myT.TextFrame.TextRange.Characters(1, 5).Font.Bold = msoTrue
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

