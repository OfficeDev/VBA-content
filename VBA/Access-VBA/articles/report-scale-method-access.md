---
title: Report.Scale Method (Access)
keywords: vbaac10.chm13785
f1_keywords:
- vbaac10.chm13785
ms.prod: access
api_name:
- Access.Report.Scale
ms.assetid: 6a261d1d-9474-7374-f399-4d46e404058b
ms.date: 06/08/2017
---


# Report.Scale Method (Access)

The  **Scale** method defines the coordinate system for a **[Report](report-object-access.md)** object.


## Syntax

 _expression_. **Scale**( ** _flags_**, ** _x1_**, ** _y1_**, ** _x2_**, ** _y2_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _flags_|Required|**Integer**||
| _x1_|Required|**Single**| A value for the horizontal coordinate that defines the position of the upper-left corner of the object.|
| _y1_|Required|**Single**| A value for the horizontal coordinate that defines the position of the upper-left corner of the object.|
| _x2_|Required|**Single**|A value for the horizontal coordinate that defines the position of the lower-right corner of the object.|
| _y2_|Required|**Single**|A value for the vertical coordinate that defines the position of the lower-right corner of the object.|

### Return Value

Nothing


## Remarks

You can use this method only in an event procedure or a macro specified by the  **OnPrint** or **OnFormat** event property for a report section, or the **OnPage** event property for a report.

You can use the  **Scale** method to reset the coordinate system to any scale you choose. Using the **Scale** method with no arguments resets the coordinate system to twips. The **Scale** method affects the coordinate system for the **Print** method and the report graphics methods, which include the **Circle**, **Line**, and **PSet** methods.


## Example

The following example draws a circle with one scale, then uses the  **Scale** method to change the scale and draw another circle with the new scale.


```vb
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
 DrawCircle 
End Sub 
 
Sub DrawCircle() 
 Dim sngHCtr As Single, sngVCtr As Single 
 Dim sngNewH As Single, sngNewV As Single 
 Dim sngRadius As Single 
 
 Me.ScaleMode = 3 ' Set scale to pixels. 
 sngHCtr = Me.ScaleWidth / 2 ' Horizontal center. 
 sngVCtr = Me.ScaleHeight / 2 ' Vertical center. 
 sngRadius = Me.ScaleHeight / 3 ' Circle radius. 
 ' Draw circle. 
 Me.Circle (sngHCtr, sngVCtr), sngRadius 
 ' New horizontal scale. 
 sngNewH = Me.ScaleWidth * 0.9 
 ' New vertical scale. 
 sngNewV = Me.ScaleHeight * 0.9 
 ' Change to new scale. 
 Me.Scale(0, 0)-(sngNewH, sngNewV) 
 ' Draw circle. 
 Me.Circle (sngHCtr + 100, sngVCtr), sngRadius, RGB(0, 256, 0) 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

