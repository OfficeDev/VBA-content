---
title: FillFormat.GradientAngle Property (Excel)
ms.prod: excel
api_name:
- Excel.FillFormat.GradientAngle
ms.assetid: cc2b2d08-1411-f79f-806c-5f832a1ce715
ms.date: 06/08/2017
---


# FillFormat.GradientAngle Property (Excel)

Returns or sets the angle of the gradient fill for the specified fill format. Read/write


## Syntax

 _expression_ . **GradientAngle**

 _expression_ A variable that represents a **[FillFormat](fillformat-object-excel.md)** object.


### Return Value

 **Single**


## Remarks

A gradient fill can be specified in the formatting for various elements (shapes) in a chart. For example, you can use the  **Format Data Series** dialog box to format the columns in a **Column** chart to a gradient fill. In this case, the **GradientAngle** property corresponds to the setting of the ** Angle** box in the **Fill** category of the **Format Data Series** dialog box. The valid range of values for the **GradientAngle** property is from 0 to 359.9.


## Example

The following code example sets the angle of the gradient fill for Series 1 in Chart 1 to 45 degrees.


```vb
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.SeriesCollection(1).Format.Fill.GradientAngle = 45
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

