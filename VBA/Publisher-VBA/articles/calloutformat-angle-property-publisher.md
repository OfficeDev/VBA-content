---
title: CalloutFormat.Angle Property (Publisher)
keywords: vbapb10.chm2490625
f1_keywords:
- vbapb10.chm2490625
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Angle
ms.assetid: b65a1c87-db52-8703-135e-1fbb1efbeebe
ms.date: 06/08/2017
---


# CalloutFormat.Angle Property (Publisher)

Returns or sets an  **MsoCalloutAngleType** constant that represents the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write.


## Syntax

 _expression_. **Angle**

 _expression_A variable that represents a  **CalloutFormat** object.


## Remarks

If you set the value of this property to anything other than  **msoCalloutAngleAutomatic**, the callout line maintains a fixed angle as you drag the callout.



|MsoCalloutAngleType can be one of these MsoCalloutAngleType constants.|
| **msoCalloutAngle30**|
| **msoCalloutAngle45**|
| **msoCalloutAngle60**|
| **msoCalloutAngle90**|
| **msoCalloutAngleAutomatic**|
| **msoCalloutAngleMixed**|

## Example

This example sets the callout angle to 90 degrees for the first shape on the first page of the active publication. For this example to work, the specified shape must be a callout.


```vb
Sub SetCalloutAngle() 
 ActiveDocument.Pages(1).Shapes(1).Callout.Angle = msoCalloutAngle90 
End Sub
```


