---
title: QuickAnalysis.Show Method (Excel)
keywords: vbaxl10.chm920073
f1_keywords:
- vbaxl10.chm920073
ms.prod: excel
ms.assetid: 0a30cfb1-1a15-95da-9ad5-2bf579696769
ms.date: 06/08/2017
---


# QuickAnalysis.Show Method (Excel)

Displays specific members of the Analysis Lens user interface.


## Syntax

 _expression_ . **Show**_(XlQuickAnalysisMode)_

 _expression_ A variable that represents a[QuickAnalysis](quickanalysis-object-excel.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XlQuickAnalysisMode_|Optional|XLQUICKANALYSISMODE||

### Remarks

When [XlQuickAnalysisMode Enumeration (Excel)](xlquickanalysismode-enumeration-excel.md) is set to any one of the following options, the resulting user interface is displayed:


- If missing or set to  **xlLensOnly** = Show the button but no fly-outs
    
-  **xlFormatConditions** = Conditional Formatting
    
-  **xlSparklines** = Sparklines
    
-  **xlRecommendedCharts** = Charts
    
-  **xlTables** = Tables
    
-  **xlTotals** = Totals
    

### Return value

 **VOID**


## See also


#### Other resources



[QuickAnalysis Object](quickanalysis-object-excel.md)

