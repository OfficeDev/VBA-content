---
title: Application.QuickAnalysis Property (Excel)
keywords: vbaxl10.chm133338
f1_keywords:
- vbaxl10.chm133338
ms.prod: excel
ms.assetid: c79c04e7-0caf-470c-ee6d-dc613d6a4cf5
ms.date: 06/08/2017
---


# Application.QuickAnalysis Property (Excel)

Returns a  **[QuickAnalysis](quickanalysis-object-excel.md)** object that represents the Quick Analysis options of the application.


## Syntax

 _expression_ . **QuickAnalysis**

 _expression_ A variable that represents an **Application** object.


## Example

The following example displays the Quick Analysis contextual UI with the  **Sparklines** option highlighted.


```vb
Sub ShowQuickAnalysisOptions()

'Displays the Quick Analysis contextual UI with the Sparklines option highlighted.
  Application.QuickAnalysis.Show (xlSparklines)

End Sub
```


## Property value

 **QUICKANALYSIS**


## See also


#### Concepts


[Application Object](application-object-excel.md)

