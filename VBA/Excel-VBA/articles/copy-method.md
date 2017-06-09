---
title: Copy Method
keywords: vbagr10.chm66087
f1_keywords:
- vbagr10.chm66087
ms.prod: excel
api_name:
- Excel.Copy
ms.assetid: 2207804d-0003-5c75-afa8-a718efba0c2c
ms.date: 06/08/2017
---


# Copy Method

Copy method as it applies to the  **ChartArea** object.

Copies a picture of the point or series to the Clipboard.

 _expression_. **Copy**

 _expression_ Required. An expression that returns one of the above objects.
Copy method as it applies to the  **Range** object.
Copies the Range to the specified range or to the Clipboard.
 _expression_. **Copy**( **_Destination_**)
 _expression_ Required. An expression that returns one of the above objects.
 **Destination** Optional **Variant**. Specifies the new range to which the specified range will be copied. If this argument is omitted, Microsoft Graph copies the range to the Clipboard.

## Example

This example copies the formulas in cells A1:D4 on the datasheet into cells E5:H8.


```vb
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:D4").Copy _ 
 Destination:= mySheet.Range("E5")
```


