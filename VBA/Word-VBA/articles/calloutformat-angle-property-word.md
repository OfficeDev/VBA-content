---
title: CalloutFormat.Angle Property (Word)
keywords: vbawd10.chm163905637
f1_keywords:
- vbawd10.chm163905637
ms.prod: word
api_name:
- Word.CalloutFormat.Angle
ms.assetid: b5178aa0-c2e3-dc59-766d-7ce5b2e7c762
ms.date: 06/08/2017
---


# CalloutFormat.Angle Property (Word)

Returns or sets the angle of the callout line. Read/write  **[MsoCalloutAngleType](http://msdn.microsoft.com/library/f4535cc0-9c8c-6579-67d5-532650dec2ef%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **Angle**

 _expression_ A variable that represents a **[CalloutFormat](calloutformat-object-word.md)** object.


## Remarks

If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. If you set the value of this property to anything other than  **msoCalloutAngleAutomatic** , the callout line maintains a fixed angle as you drag the callout.


 **Note**  Setting this property to  **msoCalloutAngleMixed** will cause an error. **msoCalloutAngleMixed** is a return value only. It indicates a combination of the other states.


## Example

This example sets the callout angle to 90 degrees for a callout named "co1" on the active document.


```vb
ActiveDocument.Shapes("co1").Callout.Angle = msoCalloutAngle90
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-word.md)

