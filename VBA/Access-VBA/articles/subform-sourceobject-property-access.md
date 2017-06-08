---
title: SubForm.SourceObject Property (Access)
keywords: vbaac10.chm11926
f1_keywords:
- vbaac10.chm11926
ms.prod: access
api_name:
- Access.SubForm.SourceObject
ms.assetid: bee9c1fe-c58c-b6f3-e2ad-7ceb99bacee4
ms.date: 06/08/2017
---


# SubForm.SourceObject Property (Access)

You can use the  **SourceObject** property to identify the form or report that is the source of the subform or subreport on a form or report. Read/write **String**.


## Syntax

 _expression_. **SourceObject**

 _expression_ A variable that represents a **SubForm** object.


## Remarks

Enter the name of the form or report that is the source of the subform or subreport in the control's property sheet. If you add a subform or subreport to the form or report by dragging it from the Database window, the  **SourceObject** property is set automatically in the property sheet.

In Visual Basic, you set this property by using a string expression that is a name of a form or report.


 **Note**  You can't set or change the  **SourceObject** property in the **Open** or **Format** events of a report.

If you delete the  **SourceObject** property setting in the property sheet for a subform or subreport, the control remains on the form but is no longer bound to the source form or report.


## See also


#### Concepts


[SubForm Object](subform-object-access.md)

