---
title: Sheets.PrintOut Method (Excel)
keywords: vbaxl10.chm152089
f1_keywords:
- vbaxl10.chm152089
ms.prod: excel
api_name:
- Excel.Sheets.PrintOut
ms.assetid: b8e11498-4a45-b0d4-9a81-779f924e4e7e
ms.date: 06/08/2017
---


# Sheets.PrintOut Method (Excel)

Prints the object.


## Syntax

 _expression_ . **PrintOut**( **_From_** , **_To_** , **_Copies_** , **_Preview_** , **_ActivePrinter_** , **_PrintToFile_** , **_Collate_** , **_PrToFileName_** , **_IgnorePrintAreas_** )

 _expression_ A variable that represents a **Sheets** object.


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


## Examples

 **Prints Active Sheet**


```vb
ActiveSheet.PrintOut
```

 **Prints from Page X to Page Y**

Prints from page 2 to page 3.




```vb
Worksheets. ("sheet1").PrintOut From:=2, To:=3
```

 **Prints 3 Copies**

Prints three copies from page 2 to page 3.




```vb
Worksheets. ("sheet1").PrintOut From:=2, To:=3, Copies:=3
```


## See also


#### Concepts


[Sheets Object](sheets-object-excel.md)

