---
title: Application.Build Property (Excel)
keywords: vbaxl10.chm133082
f1_keywords:
- vbaxl10.chm133082
ms.prod: excel
api_name:
- Excel.Application.Build
ms.assetid: da8ec8af-c1d9-917e-a057-a4762a783124
ms.date: 06/08/2017
---


# Application.Build Property (Excel)

Returns the Microsoft Excel build number. Read-only  **Long** .


## Syntax

 _expression_ . **Build**

 _expression_ A variable that represents an **Application** object.


## Remarks

It's usually safer to test the  **[Version](application-version-property-excel.md)** property, unless you're sure you need to know the build number.


## Example

This example tests the  **Build** property.


```vb
If Application.Build > 2500 Then 
 ' build-dependent code here 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

