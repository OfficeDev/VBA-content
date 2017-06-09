---
title: AddChartAutoFormat Method
keywords: vbagr10.chm65752
f1_keywords:
- vbagr10.chm65752
ms.prod: excel
api_name:
- Excel.AddChartAutoFormat
ms.assetid: a75fe1dd-72f5-526c-a9b4-a309825e84d7
ms.date: 06/08/2017
---


# AddChartAutoFormat Method

Adds a custom chart autoformat to the list of available chart autoformats.

 _expression_. **AddChartAutoFormat( _Name_**,  **_Description_)**

 _expression_ Required. An expression that returns an **Application** object.

 **Name** Required **String**. The name of the autoformat.
 **Description** Optional **String**. A description of the custom autoformat.

## Example

This example adds a new autoformat.


```
myChart.Application.AddChartAutoFormat _ 
 Name:="Presentation Chart"
```


