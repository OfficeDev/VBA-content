---
title: XlTrendlineType Enumeration (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.XlTrendlineType
ms.assetid: c0a39c74-2df1-9816-0daf-3fce481a2241
ms.date: 06/08/2017
---


# XlTrendlineType Enumeration (PowerPoint)

Specifies how the trendline that smoothes out fluctuations in the data is calculated.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**xlExponential**|5|Uses an equation to calculate the least squares fit through points (for example, y=ab^x) .|
|**xlLinear**|-4132|Uses the linear equation y = mx + b to calculate the least squares fit through points.|
|**xlLogarithmic**|-4133|Uses the equation y = c ln x + b to calculate the least squares fit through points.|
|**xlMovingAvg**|6|Uses a sequence of averages computed from parts of the data series. The number of points equals the total number of points in the series minus the number specified for the period.|
|**xlPolynomial**|3|Uses an equation to calculate the least squares fit through points (for example, y = ax^6 + bx^5 + cx^4 + dx^3 + ex^2 + fx + g).|
|**xlPower**|4|Uses an equation to calculate the least squares fit through points (for example, y = ax^b).|

