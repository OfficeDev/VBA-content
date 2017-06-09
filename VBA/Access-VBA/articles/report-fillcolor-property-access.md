---
title: Report.FillColor Property (Access)
keywords: vbaac10.chm13756
f1_keywords:
- vbaac10.chm13756
ms.prod: access
api_name:
- Access.Report.FillColor
ms.assetid: 04fa1376-fddb-a4b3-04fd-d562f0567136
ms.date: 06/08/2017
---


# Report.FillColor Property (Access)

You use the  **FillColor** property to specify the color that fills in boxes and circles drawn on reports with the **[Line](report-line-method-access.md)** and **[Circle](report-circle-method-access.md)** methods. You can also use this property with[Visual Basic](set-properties-by-using-visual-basic.md)to create special visual effects on custom reports when you print using a color printer or preview the reports on a color monitor. Read/write  **Long**.


## Syntax

 _expression_. **FillColor**

 _expression_ A variable that represents a **Report** object.


## Remarks

You can set this property only in an event procedure specified by a section's  **OnPrint** event property.

The following example uses the  **Circle** method to draw a circle and create a pie slice within the circle. Then it uses the **FillColor** and **FillStyle** properties to color the pie slice red. It also draws a line from the upper left to the center of the circle.

You can use the  **RGB** or **QBColor** functions to set this property. The **FillColor** property setting has a data type of **Long**.


## Example

To try this example in Microsoft Access, create a new report. Set the  **OnPrint** property of the Detail section to [Event Procedure]. Enter the following code in the report's module, then switch to Print Preview.


```vb
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
 
 Const conPI = 3.14159265359 
 
 Dim sngHCtr As Single 
 Dim sngVCtr As Single 
 Dim sngRadius As Single 
 Dim sngStart As Single 
 Dim sngEnd As Single 
 
 sngHCtr = Me.ScaleWidth / 2 ' Horizontal center. 
 sngVCtr = Me.ScaleHeight / 2 ' Vertical center. 
 sngRadius = Me.ScaleHeight / 3 ' Circle radius. 
 Me.Circle (sngHCtr, sngVCtr), sngRadius ' Draw circle. 
 sngStart = -0.00000001 ' Start of pie slice. 
 
 sngEnd = -2 * conPI / 3 ' End of pie slice. 
 Me.FillColor = RGB(255, 0, 0) ' Color pie slice red. 
 Me.FillStyle = 0 ' Fill pie slice. 
 
 ' Draw Pie slice within circle 
 Me.Circle (sngHCtr, sngVCtr), sngRadius, , sngStart, sngEnd 
 
 ' Draw line to center of circle. 
 Dim intColor As Integer 
 Dim sngTop As Single, sngLeft As Single 
 Dim sngWidth As Single, sngHeight As Single 
 
 Me.ScaleMode = 3 ' Set scale to pixels. 
 sngTop = Me.ScaleTop ' Top inside edge. 
 sngLeft = Me.ScaleLeft ' Left inside edge. 
 sngWidth = Me.ScaleWidth / 2 ' Width inside edge. 
 sngHeight = Me.ScaleHeight / 2 ' Height inside edge. 
 intColor = RGB(255, 0, 0) ' Make color red. 
 
 ' Draw line. 
 Me.Line (sngTop, sngLeft)-(sngWidth, sngHeight), intColor 
 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

