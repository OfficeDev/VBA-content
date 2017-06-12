---
title: Report.Top Property (Access)
keywords: vbaac10.chm13728
f1_keywords:
- vbaac10.chm13728
ms.prod: access
api_name:
- Access.Report.Top
ms.assetid: badaa1a0-44ef-c2cd-64fa-8450add21d69
ms.date: 06/08/2017
---


# Report.Top Property (Access)

You can use the  **Top** property to specify an object's location on a form or report. Read/write **Long**. .


## Syntax

 _expression_. **Top**

 _expression_ A variable that represents a **Report** object.


## Remarks

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips.

For reports, the  **Top** property setting is the amount the current section is offset from the top of the page. This property setting is expressed in twips. You can use this property to specify how far down the page you want a section to print in the section's **Format** event procedure.


## Example

The following example checks the  **Top** property setting for the current report. If the value is less than the minimum margin setting, the **NextRecord** and **PrintSection** properties are set to **False**. The section doesn't advance to the next record, and the next section isn't printed.


```vb
Sub Detail1_Format(Cancel As Integer, FormatCount As Integer) 
Const conTopMargin = 1880 
' Don't advance to next record or print next section 
' if Top property setting is less than 1880 twips. 
 If Me.Top < conTopMargin Then 
 Me.NextRecord = False 
 Me.PrintSection = False 
 End If 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

