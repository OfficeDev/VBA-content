---
title: Report.Modal Property (Access)
keywords: vbaac10.chm13799
f1_keywords:
- vbaac10.chm13799
ms.prod: access
api_name:
- Access.Report.Modal
ms.assetid: 654ff830-c8d9-5bd9-1ec6-61ee6546b4db
ms.date: 06/08/2017
---


# Report.Modal Property (Access)

You can use the  **Modal** property to specify whether a report opens as a modal window. When a report opens as a modal window, you must close the window before you can move the focus to another object. Read/write **Boolean**.


## Syntax

 _expression_. **Modal**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **Modal** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The form or report opens as a modal window.|
|No|**False**|(Default) The form opens as a non-modal window.|
When you open a modal window, other windows in Microsoft Access are disabled until you close it (although you can switch to windows in other applications). To disable menus and toolbars in addition to other windows, set both the  **Modal** and **PopUp** properties to Yes.


## See also


#### Concepts


[Report Object](report-object-access.md)

