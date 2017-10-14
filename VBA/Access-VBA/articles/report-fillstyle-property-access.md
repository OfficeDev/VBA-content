---
title: Report.FillStyle Property (Access)
keywords: vbaac10.chm13757
f1_keywords:
- vbaac10.chm13757
ms.prod: access
api_name:
- Access.Report.FillStyle
ms.assetid: 0fcb840d-4ff6-718a-2267-25cd2622c8d2
ms.date: 06/08/2017
---


# Report.FillStyle Property (Access)

You can use the  **FillStyle** property to specify whether a circle or line drawn by the **[Circle](report-circle-method-access.md)** or **[Line](report-line-method-access.md)** method on a report is transparent, opaque, or filled with a pattern. Read/write **Integer**.


## Syntax

 _expression_. **FillStyle**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **FillStyle** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|Opaque|
|1|(Default) Transparent|
|2|Horizontal Line|
|3|Vertical Line|
|4|Upward Diagonal|
|5|Downward Diagonal|
|6|Cross|
|7|Diagonal Cross|

 **Note**  You can set the  **FillStyle** property in an event procedure specified by a section's **OnPrint** property setting.

When the  **FillStyle** property is set to 0, a circle or line has the color set by the **[FillColor](report-fillcolor-property-access.md)** property. When the **FillStyle** property is set to 1, the interior of the circle or line is transparent and has the color of the report behind it.

To use the  **FillStyle** property, the **[SpecialEffect](line-specialeffect-property-access.md)** property must be set to Normal.

The following example uses the  **Circle** method to draw a circle and create a pie slice within the circle. Then it uses the **FillColor** and **FillStyle** properties to color the pie slice red. It also draws a line from the upper left to the center of the circle.


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

