---
title: CalloutFormat.Angle Property (Excel)
keywords: vbaxl10.chm104007
f1_keywords:
- vbaxl10.chm104007
ms.prod: excel
api_name:
- Excel.CalloutFormat.Angle
ms.assetid: 8f3dab54-4597-e22c-ae3e-cf894849b668
ms.date: 06/08/2017
---


# CalloutFormat.Angle Property (Excel)

Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write  **[MsoCalloutAngleType](http://msdn.microsoft.com/library/f4535cc0-9c8c-6579-67d5-532650dec2ef%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **Angle**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks

If you set the value of this property to anything other than  **msoCalloutAngleAutomatic** , the callout line maintains a fixed angle as you drag the callout.


## Example

This example sets to 90 degrees the callout angle for a callout named "callout1" on  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes("callout1").Callout.Angle = msoCalloutAngle90
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-excel.md)

