---
title: Report.Height Property (Access)
keywords: vbaac10.chm13729
f1_keywords:
- vbaac10.chm13729
ms.prod: access
api_name:
- Access.Report.Height
ms.assetid: 14821735-efbb-e831-e1d4-94f34de41ef7
ms.date: 06/08/2017
---


# Report.Height Property (Access)

Gets or sets the height of the specified object in twips. Read/write  **Long**.


## Syntax

 _expression_. **Height**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **Height** property applies only to form sections and report sections, not to forms and reports.

For report sections, you can't use a macro or Visual Basic to set the  **Height** property when you print or preview a report. For report controls, you can set the **Height** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

You can't set this propery for an object once the print process has started. For example, attempting to set the  **Height** property in a report's **Print** event generates an error.

Microsoft Access automatically sets the  **Height** property when you create or size a control or when you size a window in form Design View or report Design view.

The height of sections is measured from the inside of their borders. The height of controls is measured from the center of their borders so controls with different border widths align correctly. The margins for forms and reports are set in the  **Page Setup** dialog box, available by clicking **Page Setup** on the **File** menu.


## See also


#### Concepts


[Report Object](report-object-access.md)

