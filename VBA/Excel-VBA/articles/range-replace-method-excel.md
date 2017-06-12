---
title: Range.Replace Method (Excel)
keywords: vbaxl10.chm144186
f1_keywords:
- vbaxl10.chm144186
ms.prod: excel
api_name:
- Excel.Range.Replace
ms.assetid: 12647334-f911-69e4-de31-b4df2722eff3
ms.date: 06/08/2017
---


# Range.Replace Method (Excel)

Returns a  **Boolean** indicating characters in cells within the specified range. Using this method doesn't change either the selection or the active cell.


## Syntax

 _expression_ . **Replace**( **_What_** , **_Replacement_** , **_LookAt_** , **_SearchOrder_** , **_MatchCase_** , **_MatchByte_** , **_SearchFormat_** , **_ReplaceFormat_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _What_|Required| **Variant**|The string you want Microsoft Excel to search for.|
| _Replacement_|Required| **Variant**|The replacement string.|
| _LookAt_|Optional| **Variant**|Can be one of the following  **[XlLookAt](xllookat-enumeration-excel.md)** constants: **xlWhole** or **xlPart** .|
| _SearchOrder_|Optional| **Variant**|Can be one of the following  **[XlSearchOrder](xlsearchorder-enumeration-excel.md)** constants: **xlByRows** or **xlByColumns** .|
| _MatchCase_|Optional| **Variant**| **True** to make the search case sensitive.|
| _MatchByte_|Optional| **Variant**|You can use this argument only if you?ve selected or installed double-byte language support in Microsoft Excel.  **True** to have double-byte characters match only double-byte characters. **False** to have double-byte characters match their single-byte equivalents.|
| _SearchFormat_|Optional| **Variant**|The search format for the method.|
| _ReplaceFormat_|Optional| **Variant**|The replace format for the method.|

### Return Value

Boolean


## Remarks

The settings for  _LookAt_,  _SearchOrder_,  _MatchCase_, and  _MatchByte_ are saved each time you use this method. If you don?t specify values for these arguments the next time you call the method, the saved values are used. Setting these arguments changes the settings in the **Find** dialog box, and changing the settings in the **Find** dialog box changes the saved values that are used if you omit the arguments. To avoid problems, set these arguments explicitly each time you use this method.


## Example

This example replaces every occurrence of the trigonometric function SIN with the function COS. The replacement range is column A on Sheet1.


```vb
Worksheets("Sheet1").Columns("A").Replace _ 
 What:="SIN", Replacement:="COS", _ 
 SearchOrder:=xlByColumns, MatchCase:=True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

