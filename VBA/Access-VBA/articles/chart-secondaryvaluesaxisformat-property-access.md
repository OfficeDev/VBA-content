---
title: Chart.SecondaryValuesAxisFormat Property (Access)
keywords: vbaac10.chm6165
f1_keywords:
- vbaac10.chm6165
ms.prod: access
api_name:
- Access.Chart.SecondaryValuesAxisFormat
ms.date: 05/02/2018
---


# Chart.SecondaryValuesAxisFormat Property (Access)

Returns or sets the format of the values on the secondary values axis. Read/write **String** .

You can use a [predefined or custom format](format-propertynumber-and-currency-data-types.md).


## Syntax

 _expression_ . **SecondaryValuesAxisFormat**

 _expression_ A variable that represents a **Chart** object.


## Example

```vb
With myChart
 .PrimaryValuesAxisFormat = "#,###.#0"
 .SecondaryValuesAxisFormat = "Currency"
End With
```

## See also


#### Concepts


[Format Property - Number and Currency Data Types](format-propertynumber-and-currency-data-types.md)

[Chart Object](chart-object-access.md)