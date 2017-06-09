---
title: Application.Union Method (Excel)
keywords: vbaxl10.chm132112
f1_keywords:
- vbaxl10.chm132112
ms.prod: excel
api_name:
- Excel.Application.Union
ms.assetid: 7c70a5be-2696-5fc2-bd69-6c6ff4d3291e
ms.date: 06/08/2017
---


# Application.Union Method (Excel)

Returns the union of two or more ranges.


## Syntax

 _expression_ . **Union**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|At least two  **[Range](range-object-excel.md)** objects must be specified.|
| _Arg2_|Required| **Range**|At least two  **Range** objects must be specified.|
| _Arg3_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg4_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg5_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg6_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg7_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg8_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg9_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg10_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg11_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg12_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg13_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg14_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg15_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg16_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg17_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg18_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg19_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg20_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg21_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg22_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg23_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg24_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg25_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg26_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg27_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg28_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg29_|Optional| **Variant**|At least two  **Range** objects must be specified.|
| _Arg30_|Optional| **Variant**|At least two  **Range** objects must be specified.|

### Return Value

Range


## Example

This example fills the union of two named ranges, Range1 and Range2, with the formula =RAND().


```vb
Worksheets("Sheet1").Activate 
Set bigRange = Application.Union(Range("Range1"), Range("Range2")) 
bigRange.Formula = "=RAND()"
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

