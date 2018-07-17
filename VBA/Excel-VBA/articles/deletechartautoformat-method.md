---
title: DeleteChartAutoFormat Method
keywords: vbagr10.chm65753
f1_keywords:
- vbagr10.chm65753
ms.prod: excel
api_name:
- Excel.DeleteChartAutoFormat
ms.assetid: 22f9c561-b0a1-2c75-391e-25bb54ad67a5
ms.date: 06/08/2017
---


# DeleteChartAutoFormat Method

Removes a custom chart autoformat from the list of available chart autoformats.

 _expression_. **DeleteChartAutoFormat( _Name_)**

 _expression_ Required. An expression that returns an **Application** object.

 **Name** Required **String**. The name of the custom autoformat to be removed.

## Example

This example deletes the custom autoformat named "Presentation Chart."


```
myChart.Application.DeleteChartAutoFormat _ 
 name:="Presentation Chart" 

```


