---
title: Report.PSet Method (Access)
keywords: vbaac10.chm13784
f1_keywords:
- vbaac10.chm13784
ms.prod: access
api_name:
- Access.Report.PSet
ms.assetid: 951a262b-b17b-9b95-b5f2-922d4aff9ce9
ms.date: 06/08/2017
---


# Report.PSet Method (Access)

The  **PSet** method sets a point on a **[Report](report-object-access.md)** object to a specified color when the **Print** event occurs.


## Syntax

 _expression_. **PSet**( ** _flags_**, ** _X_**, ** _Y_**, ** _color_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _flags_|Required|**Integer**| A keyword that indicates the coordinates are relative to the current graphics position given by the settings for the **[CurrentX](report-currentx-property-access.md)** and **[CurrentY](report-currenty-property-access.md)** properties of the _object_ argument.|
| _X_|Required|**Single**|The horizontal coordinate of the point to set.|
| _Y_|Required|**Single**|The vertical coordinate of the point to set.|
| _color_|Required|**Long**|the RGB (red-green-blue) color to set the point to. If this argument is omitted, the value of the  **ForeColor** property is used. You can also use the **RGB** function or **QBColor** function to specify the color.|

### Return Value

Nothing


## Remarks

The size of the point depends on the  **[DrawWidth](report-drawwidth-property-access.md)** property setting. When the **DrawWidth** property is set to 1, the **PSet** method sets a single pixel to the specified color. When the **DrawWidth** property is greater than 1, the point is centered on the specified coordinates.

The way the point is drawn depends on the settings of the  **[DrawMode](report-drawmode-property-access.md)** and **[DrawStyle](report-drawstyle-property-access.md)** properties.

When you apply the  **PSet** method, the **CurrentX** and **CurrentY** properties are set to the point specified by the _x_ and _y_ arguments.

To clear a single pixel with the  **PSet** method, specify the coordinates of the pixel and use &;HFFFFFF (white) as the _color_ argument.


## Example

The following example uses the  **PSet** method to draw a line through the horizontal axis of a report.

To try this example in Microsoft Access, create a new report. Set the  **OnPrint** property of the Detail section to [Event Procedure]. Enter the following code in the report's module, then switch to Print Preview.




```vb
Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
 Dim sngMidPt As Single, intI As Integer 
 ' Set scale to pixels. 
 Me.ScaleMode = 3 
 ' Calculate midpoint. 
 sngMidPt = Me.ScaleHeight / 2 
 ' Loop to draw line down horizontal axis pixel by pixel. 
 For intI = 1 To Me.ScaleWidth 
 Me.PSet(intI, sngMidPt) 
 Next intI 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

