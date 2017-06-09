---
title: Report.Line Method (Access)
keywords: vbaac10.chm13783
f1_keywords:
- vbaac10.chm13783
ms.prod: access
api_name:
- Access.Report.Line
ms.assetid: 9e640e37-c055-3dc3-b70e-0805cdc13561
ms.date: 06/08/2017
---


# Report.Line Method (Access)

The  **Line** method draws lines and rectangles on a **Report** object when the Print event occurs.


## Syntax

 _expression_. **Line** ( ** _Step_** ( **x1,  _y1_** ) - ** _Step_** ( ** _x2, y2_** ), ** _color_**, ** _BF_** )

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


### Parameters

 _Step_ A keyword that indicates the starting point coordinates are relative to the current graphics position given by the current settings for the **CurrentX** and **CurrentY** properties of the object argument.

 _object_ **x1, y1** **Single** values indicating the coordinates of the starting point for the line or rectangle. The Scale properties ( **[ScaleMode](report-scalemode-property-access.md)**, **ScaleLeft**, **ScaleTop**, **ScaleHeight**, and **ScaleWidth** ) of the **Report** object specified by the _object_ argument determine the unit of measure used. If this argument is omitted, the line begins at the position indicated by the **CurrentX** and **CyrrentY** properties.

 _color_ A **Long** value indicating the RGB (red-green-blue) color used to draw the line. If this argument is omitted, the value of the **ForeColor** property is used. You can also use the **RGB** function or **QBColor** function to specify the color.

 _B_ An option that creates a rectangle by using the coordinates as opposite corners of the rectangle.

 _F_ **F** cannot be used without **B**. If the **B** option is used, the **F** option specifies that the rectangle is filled with the same color used to draw the rectangle. If **B** is used without **F**, the rectangle is filled with the color specified by the current settings of the **FillColor** and **BackStyle** properties. The default value for the **BackStyle** property is Normal for rectangles and lines.


## Remarks

You can use this method only in an event procedure or a macro specified by the  **OnPrint** or **OnFormat** event property for a report section, or the **OnPage** event property for a report.

To connect two drawing lines, make sure that one line begins at the end point of the previous line.

The width of the line drawn depends on the  **DrawWidth** property setting. The way a line or rectangle is drawn on the background depends on the settings of the **DrawMode** and **DrawStyle** properties.

When you apply the  **Line** method, the **CurrentX** and **CurrentY** properties are set to the end point specified by the _x2_ and _y2_ arguments.


## Example

The following example uses the  **Line** method to draw a red rectangle five pixels inside the edge of a report named EmployeeReport. The **RGB** function is used to make the line red.

To try this example in Microsoft Access, create a new report named EmployeeReport. Paste the following code in the declarations section of the report's module, then switch to Print Preview.




```vb
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
    ' Call the Drawline procedure 
    DrawLine 
End Sub 
 
Sub DrawLine() 
    Dim rpt As Report, lngColor As Long 
    Dim sngTop As Single, sngLeft As Single 
    Dim sngWidth As Single, sngHeight As Single 
 
    Set rpt = Reports!EmployeeReport 
    ' Set scale to pixels. 
    rpt.ScaleMode = 3 
    ' Top inside edge. 
    sngTop = rpt.ScaleTop + 5 
    ' Left inside edge. 
    sngLeft = rpt.ScaleLeft + 5 
    ' Width inside edge. 
    sngWidth = rpt.ScaleWidth - 10 
    ' Height inside edge. 
    sngHeight = rpt.ScaleHeight - 10 
    ' Make color red. 
    lngColor = RGB(255,0,0) 
    ' Draw line as a box. 
    rpt.Line(sngTop, sngLeft) - (sngWidth, sngHeight), lngColor, B 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

