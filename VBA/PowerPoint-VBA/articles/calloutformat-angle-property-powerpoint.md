---
title: CalloutFormat.Angle Property (PowerPoint)
keywords: vbapp10.chm559007
f1_keywords:
- vbapp10.chm559007
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.Angle
ms.assetid: 75ce8b84-f7f5-a15a-291b-3f9713bddce7
ms.date: 06/08/2017
---


# CalloutFormat.Angle Property (PowerPoint)

Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write.


## Syntax

 _expression_. **Angle**

 _expression_ A variable that represents an **CalloutFormat** object.


### Return Value

MsoCalloutAngleType


## Remarks

The value of the  **Angle** property can be one of these **MsoCalloutAngleType** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoCalloutAngle30**||
|**msoCalloutAngle45**||
|**msoCalloutAngle60**||
|**msoCalloutAngle90**||
|**msoCalloutAngleAutomatic**|The callout line maintains a fixed angle as you drag the callout.|
|**msoCalloutAngleMixed**||

## Example

This example sets to 90 degrees the callout angle for a callout named "co1" on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes("co1").Callout.Angle = msoCalloutAngle90
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

