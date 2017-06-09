---
title: Application.Intersect Method (Excel)
keywords: vbaxl10.chm183099
f1_keywords:
- vbaxl10.chm183099
ms.prod: excel
api_name:
- Excel.Application.Intersect
ms.assetid: 856d052a-3207-ced2-941c-b466cb880a93
ms.date: 06/08/2017
---


# Application.Intersect Method (Excel)

Returns a [Range](range-object-excel.md) object that represents the rectangular intersection of two or more ranges. If one or more ranges from a different worksheet are specified, an error will be returned.


## Syntax

 _expression_ . **Intersect**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg2_|Required| **Range**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg3_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg4_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg5_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg6_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg7_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg8_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg9_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg10_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg11_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg12_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg13_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg14_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg15_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg16_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg17_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg18_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg19_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg20_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg21_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg22_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg23_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg24_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg25_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg26_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg27_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg28_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg29_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|
| _Arg30_|Optional| **Variant**|The intersecting ranges. At least two  **Range** objects must be specified.|

### Return Value

Range


## Example

This example selects the intersection of two named ranges, rg1 and rg2, on Sheet1. If the ranges don't intersect, the example displays a message.


```vb
Worksheets("Sheet1").Activate 
Set isect = Application.Intersect(Range("rg1"), Range("rg2")) 
If isect Is Nothing Then 
 MsgBox "Ranges do not intersect" 
Else 
 isect.Select 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

