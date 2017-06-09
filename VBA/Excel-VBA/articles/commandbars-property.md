---
title: CommandBars Property
keywords: vbagr10.chm66975
f1_keywords:
- vbagr10.chm66975
ms.prod: excel
api_name:
- Excel.CommandBars
ms.assetid: 70c5ec17-7ce0-fd21-4f2f-6601b189266e
ms.date: 06/08/2017
---


# CommandBars Property

Returns a CommandBars object that represents the Microsoft Graph command bars. Read-only CommandBars object.

 _expression_. **CommandBars**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example deletes all custom command bars that aren't visible.


```vb
For Each bar In myChart.Application.CommandBars 
 If Not bar.BuiltIn And Not bar.Visible Then bar.Delete 
Next
```


