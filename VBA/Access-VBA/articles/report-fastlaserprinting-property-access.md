---
title: Report.FastLaserPrinting Property (Access)
keywords: vbaac10.chm13716
f1_keywords:
- vbaac10.chm13716
ms.prod: access
api_name:
- Access.Report.FastLaserPrinting
ms.assetid: b96ec618-de46-8802-0d9e-064fd8835fbd
ms.date: 06/08/2017
---


# Report.FastLaserPrinting Property (Access)

You can use the  **FastLaserPrinting** property to specify whether lines and rectangles are replaced by text character lines — similar to the underscore (_) and vertical bar (|) characters — when you print a report using most laser printers. Replacing lines and rectangles with text character lines can make printing much faster. Read/write **Boolean**.


## Syntax

 _expression_. **FastLaserPrinting**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **FastLaserPrinting** property affects any line or rectangle on a form or report, including controls that have these shapes (for example, a border around a text box).

This property has no effect on PostScript printers, dot-matrix printers, or earlier versions of laser printers that don't support text character lines.

When this property is set to  **True** and the form or report being printed has overlapping rectangles or lines, the rectangles or lines on top don't erase the rectangles or lines they overlap. If you require overlapping graphics on your report, set the **FastLaserPrinting** property to **False**.


## Example

The following example shows how to set the  **FastLaserPrinting** property for the Invoice report while in report Design view:


```vb
DoCmd.OpenReport "Invoice", acDesign 
Reports!Invoice.FastLaserPrinting = True 
DoCmd.Close acReport, "Invoice", acSaveYes
```


## See also


#### Concepts


[Report Object](report-object-access.md)

