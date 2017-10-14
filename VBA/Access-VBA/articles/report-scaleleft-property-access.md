---
title: Report.ScaleLeft Property (Access)
keywords: vbaac10.chm13744
f1_keywords:
- vbaac10.chm13744
ms.prod: access
api_name:
- Access.Report.ScaleLeft
ms.assetid: 1e20b9ca-5b5b-2b05-431e-1957f5c70524
ms.date: 06/08/2017
---


# Report.ScaleLeft Property (Access)

You can use the  **ScaleLeft** property to specify the units for the horizontal coordinates that describe the location of the left edge of a page when the **[Circle](report-circle-method-access.md)**, **[Line](report-line-method-access.md)**, **[Pset](report-pset-method-access.md)**, or **[Print](report-print-method-access.md)** method is used while a report is previewed, printed, or its output is saved to a file. Read / write **Single**.


## Syntax

 _expression_. **ScaleLeft**

 _expression_ A variable that represents a **Report** object.


## Remarks

You can set the  **ScaleLeft** property by using a macro or a[Visual Basic](set-properties-by-using-visual-basic.md)event procedure specified by a section's **OnPrint** property setting.

By using these properties and the related  **ScaleHeight** and **ScaleWidth** properties, you can set up a custom coordinate system with both positive and negative coordinates. All four of these Scale properties interact with the **[ScaleMode](report-scalemode-property-access.md)** property in the following ways:


- Setting any other Scale property to any value automatically sets the  **ScaleMode** property to 0.
    
- Setting the  **ScaleMode** property to a number greater than 0 changes the **ScaleHeight** and **ScaleWidth** property settings to the new unit of measurement and sets the **ScaleLeft** and **ScaleTop** properties to 0. Also, the **[CurrentX](report-currentx-property-access.md)** and **[CurrentY](report-currenty-property-access.md)** property settings change to reflect the new coordinates of the current point.
    
You can also use the  **Scale** method to set the **ScaleHeight**, **ScaleWidth**, **ScaleLeft**, and **ScaleTop** properties in one statement.


 **Note**  The  **ScaleLeft** property isn't the same as the **Left** property.


## Example

The following example uses the  **Circle** method to draw a circle and create a pie slice within the circle. Then it uses the **FillColor** and **FillStyle** properties to color the pie slice red. It also draws a line from the upper left to the center of the circle.

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

