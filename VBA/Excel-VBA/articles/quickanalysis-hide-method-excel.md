---
title: QuickAnalysis.Hide Method (Excel)
keywords: vbaxl10.chm920074
f1_keywords:
- vbaxl10.chm920074
ms.prod: excel
ms.assetid: dc3b805a-8744-1f63-0509-32b8254958b8
ms.date: 06/08/2017
---


# QuickAnalysis.Hide Method (Excel)

Hides specific members of the Analysis Lens user interface.


## Syntax

 _expression_ . **Hide**_(XlQuickAnalysisMode)_

 _expression_ A variable that represents a[QuickAnalysis](quickanalysis-object-excel.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XlQuickAnalysisMode_|Optional|XLQUICKANALYSISMODE||

### Return value

 **VOID**


## Remarks

When the argument is set to any one of the following options, the resulting user interface is hidden:


- If missing or set to  **0** = Hide all buttons
    
-  **1** = If showing, hide the **Conditional Formatting** &; **Sparklines** buttons
    
-  **2** = If showing, hide the **Charts** button
    
-  **3** = If showing, hide **Suggested Views** button
    
-  **4** = If showing, hide the **Totals** button
    

## See also


#### Other resources



[QuickAnalysis Object](quickanalysis-object-excel.md)

