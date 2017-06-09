---
title: Report.ScrollBars Property (Access)
keywords: vbaac10.chm13820
f1_keywords:
- vbaac10.chm13820
ms.prod: access
api_name:
- Access.Report.ScrollBars
ms.assetid: 12693642-6288-4f21-40cd-5aa1d6886cca
ms.date: 06/08/2017
---


# Report.ScrollBars Property (Access)

Gets or sets whether scroll bars appear on a report. Read/write  **Byte**.


## Syntax

 _expression_. **ScrollBars**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **ScrollBars** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Neither |0|No scroll bars appear on the report.|
|Horizontal Only |1|Horizontal scroll bar appears on the report.|
|Vertical Only |2|Vertical scroll bar appears on the report.|
|Both|3|(Default) Vertical and horizontal scroll bars appear on the report. |
If your report is larger than the available display window, you can use the  **ScrollBars** property to allow the user to view the entire report.


## See also


#### Concepts


[Report Object](report-object-access.md)

