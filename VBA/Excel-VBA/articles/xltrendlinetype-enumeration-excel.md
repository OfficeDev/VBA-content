---
title: XlTrendlineType Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlTrendlineType
ms.assetid: 00a95fb4-00f7-7630-2298-ef3bd7331f89
ms.date: 06/08/2017
---


# XlTrendlineType Enumeration (Excel)

Specifies how the trendline that smoothes out fluctuations in the data is calculated.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlExponential**|5|Uses an equation to calculate the least squares fit through points, for example, y=ab^x .|
| **xlLinear**|-4132|Uses the linear equation y = mx + b to calculate the least squares fit through points.|
| **xlLogarithmic**|-4133|Uses the equation y = c ln x + b to calculate the least squares fit through points.|
| **xlMovingAvg**|6|Uses a sequence of averages computed from parts of the data series. The number of points equals the total number of points in the series less the number specified for the period.|
| **xlPolynomial**|3|Uses an equation to calculate the least squares fit through points, for example, y = ax^6 + bx^5 + cx^4 + dx^3 + ex^2 + fx + g.|
| **xlPower**|4|Uses an equation to calculate the least squares fit through points, for example, y = ax^b.|

