---
title: Charts.PrintOut Method (Excel)
keywords: vbaxl10.chm218075
f1_keywords:
- vbaxl10.chm218075
ms.prod: excel
api_name:
- Excel.Charts.PrintOut
ms.assetid: ad6e659e-0fa8-a0c0-1a24-a0ec0e3b55b8
ms.date: 06/08/2017
---


# Charts.PrintOut Method (Excel)

Prints the object.


## Syntax

 _expression_ . **PrintOut**( **_From_** , **_To_** , **_Copies_** , **_Preview_** , **_ActivePrinter_** , **_PrintToFile_** , **_Collate_** , **_PrToFileName_** , **_IgnorePrintAreas_** )

 _expression_ A variable that represents a **Charts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _From_|Optional| **Variant**|The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning.|
| _To_|Optional| **Variant**|The number of the last page to print. If this argument is omitted, printing ends with the last page.|
| _Copies_|Optional| **Variant**|The number of copies to print. If this argument is omitted, one copy is printed.|
| _Preview_|Optional| **Variant**| **True** to have Microsoft Excel invoke print preview before printing the object. **False** (or omitted) to print the object immediately.|
| _ActivePrinter_|Optional| **Variant**|Sets the name of the active printer.|
| _PrintToFile_|Optional| **Variant**| **True** to print to a file. If _PrToFileName_ is not specified, Microsoft Excel prompts the user to enter the name of the output file.|
| _Collate_|Optional| **Variant**| **True** to collate multiple copies.|
| _PrToFileName_|Optional| **Variant**|If  _PrintToFile_ is set to **True** , this argument specifies the name of the file you want to print to.|
| _IgnorePrintAreas_|Optional| **Variant**| **True** to ignore print areas and print the entire object.|

### Return Value

Variant


## Remarks

"Pages" in the descriptions of  _From_ and _To_ refers to printed pages ? not overall pages in the sheet or workbook.


## Example

This example prints the active sheet.


```vb
ActiveSheet.PrintOut
```


## See also


#### Concepts


[Charts Collection](charts-object-excel.md)

