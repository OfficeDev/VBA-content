---
title: Trendline.InterceptIsAuto Property (PowerPoint)
keywords: vbapp10.chm65723
f1_keywords:
- vbapp10.chm65723
ms.prod: powerpoint
api_name:
- PowerPoint.Trendline.InterceptIsAuto
ms.assetid: 568c57e5-c42f-8559-9c7c-30a72e46463a
ms.date: 06/08/2017
---


# Trendline.InterceptIsAuto Property (PowerPoint)

 **True** if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write **Boolean**.


## Syntax

 _expression_. **InterceptIsAuto**

 _expression_ A variable that represents a **[Trendline](trendline-object-powerpoint.md)** object.


## Remarks

Setting the  **[Intercept](trendline-intercept-property-powerpoint.md)** property sets this property to **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets Microsoft Word to automatically determine the trendline intercept point for the first chart in the active document. You should run the example on a 2-D column chart that contains a single series that has a trendline.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).Trendlines(1) _
            .InterceptIsAuto = True
    End If
End With
```


## See also


#### Concepts


[Trendline Object](trendline-object-powerpoint.md)

