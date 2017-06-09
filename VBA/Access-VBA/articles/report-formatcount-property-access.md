---
title: Report.FormatCount Property (Access)
keywords: vbaac10.chm13733
f1_keywords:
- vbaac10.chm13733
ms.prod: access
api_name:
- Access.Report.FormatCount
ms.assetid: 35fbc0fb-a106-11d6-26db-99d6f0b969a3
ms.date: 06/08/2017
---


# Report.FormatCount Property (Access)

You can use the  **FormatCount** property to determine the number of times the **[OnFormat](section-onformat-property-access.md)** property has been evaluated for the current section on a report. Read/write **Integer**.


## Syntax

 _expression_. **FormatCount**

 _expression_ A variable that represents a **Report** object.


## Remarks

You can use this property only in an [event procedure](set-properties-by-using-visual-basic.md)specified by a section's  **OnFormat** property setting.

This property isn't available in report Design view.

Microsoft Access increments the  **FormatCount** property each time the **OnFormat** property setting is evaluated for the current section. As the next section is formatted, Microsoft Access resets the **FormatCount** property to 1.

Under some circumstances, Microsoft Access formats a section more than once. For example, you might design a report in which the  **[KeepTogether](section-keeptogether-property-access.md)** property for the detail section is set to Yes. When Microsoft Access reaches the bottom of a page, it formats the current detail section once to see if it will fit. If it doesn't fit, Microsoft Access moves to the next page and formats the detail section again. In this case, the setting for the **FormatCount** property for the detail section is 2 because it was formatted twice before it was printed.

You can use the  **FormatCount** property to ensure that an operation that affects formatting gets executed only once for a section.


## Example

In the following example, the  **DLookup** function is evaluated only when the **FormatCount** property is set to 1:


```vb
Private Sub Detail_Format(Cancel As Integer, _ 
 FormatCount As Integer) 
 Const conBold = 700 
 Const conNormal = 400 
 If FormatCount = 1 Then 
 If DLookup("CompanyName", _ 
 "Customers", "CustomerID = Reports!" _ 
 &; "[Customer Labels]!CustomerID") _ 
 Like "B*" Then 
 CompanyNameLine.FontWeight = conBold 
 Else 
 CompanyNameLine.FontWeight = conNormal 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

