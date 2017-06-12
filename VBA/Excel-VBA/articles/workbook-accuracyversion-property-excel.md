---
title: Workbook.AccuracyVersion Property (Excel)
keywords: vbaxl10.chm199271
f1_keywords:
- vbaxl10.chm199271
ms.prod: excel
api_name:
- Excel.Workbook.AccuracyVersion
ms.assetid: bc81782c-662c-87ec-8381-d06e77674d0c
ms.date: 06/08/2017
---


# Workbook.AccuracyVersion Property (Excel)

Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Read/write


## Syntax

 _expression_ . **AccuracyVersion**

 _expression_ A variable that represents a **[Workbook](workbook-object-excel.md)** object.


### Return Value

 **Integer**


## Remarks

By default, some of the worksheet functions from Excel 2007 and earlier versions of Excel use new algorithms that increase their accuracy. However, in some cases, the new algorithms decrease the performance of these functions relative to their performance in Excel 2007 and earlier versions of Excel. To specify that these worksheet functions use the older algorithms to increase their performance, set the  **AccuracyVersion** property to 1. The following table describes the possible settings for the **AccuracyVersion** property.



|**Setting**|**Description**|
|:-----|:-----|
|0|Use the latest, most accurate algorithms (default)|
|1|Use Excel 2007 or earlier version algorithms|
|2|Use Excel 2010 algorithms|

 **Note**  Setting the  **AccuracyVersion** property to a value other than 0, 1, or 2 will result in undefined behavior.

The following table summarizes which functions are affected by setting the  **AccuracyVersion** property to 1. The functions that are listed in the "Functions not affected" column will always use the latest accuracy algorithms or were not changed in Excel 2010. For function names that include a period (.) in their names, to determine the name of the corresponding function implemented in VBA as a method of the **[WorksheetFunction](worksheetfunction-object-excel.md)** object, substitute the underscore character ( _ ) for the period. For example, the VBA method that corresponds to the BETA.DIST function is the **[Beta_Dist](worksheetfunction-beta_dist-method-excel.md)** method.


## Example

The following example sets the affected worksheet functions to use the older algorithms to calculate their results.


```vb
ActiveWorkbook.AccuracyVersion = 1
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

