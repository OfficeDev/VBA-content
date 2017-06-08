---
title: Axis.TickMarkSpacing Property (PowerPoint)
keywords: vbapp10.chm682031
f1_keywords:
- vbapp10.chm682031
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.TickMarkSpacing
ms.assetid: 85c37d23-b91a-b390-4475-a4afa21d1566
ms.date: 06/08/2017
---


# Axis.TickMarkSpacing Property (PowerPoint)

Returns or sets the number of categories or series between tick marks. Read/write  **Long**.


## Syntax

 _expression_. **TickMarkSpacing**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

This property applies only to category and series axes. It can be a value from 1 through 31999. 

Use the  **[MajorUnit](axis-majorunit-property-powerpoint.md)** and **[MinorUnit](axis-minorunit-property-powerpoint.md)** properties to set tick-mark spacing on the value axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the number of categories between tick marks on the category axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlCategory).TickMarkSpacing = 10

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

