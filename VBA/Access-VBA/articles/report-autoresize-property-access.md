---
title: Report.AutoResize Property (Access)
keywords: vbaac10.chm13796
f1_keywords:
- vbaac10.chm13796
ms.prod: access
api_name:
- Access.Report.AutoResize
ms.assetid: bf18b1b2-aba6-d4fe-7916-de821c76fbb4
ms.date: 06/08/2017
---


# Report.AutoResize Property (Access)

Returns or sets a  **Boolean** indicating whether a Report window opens automatically sized to display complete records. Read/write.


## Syntax

 _expression_. **AutoResize**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **AutoResize** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|(Default) The Report window is automatically sized to display a complete record.|
|No|**False**|When opened, the Report window has the last saved size. To save a window's size, open the form, size the window, save the form by clicking  **Save** on the **File** menu, and close the form or report. When you next open the form or report, it will be the saved window size.|
This property can be set only in Design view.


## See also


#### Concepts


[Report Object](report-object-access.md)

