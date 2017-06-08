---
title: Chart Object Events
keywords: vbaxl10.chm5199755
f1_keywords:
- vbaxl10.chm5199755
ms.prod: excel
ms.assetid: 6808dfde-94d0-afb0-b245-44d8d1d6241e
ms.date: 06/08/2017
---


# Chart Object Events

Chart events occur when the user activates or changes a chart. Events on chart sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and select  **View Code** from the shortcut menu. Select the event name from the **Procedure** drop-down list box.

[Activate](chart-activate-event-excel.md) | 
[BeforeDoubleClick](chart-beforedoubleclick-event-excel.md) | 
[BeforeRightClick](chart-beforerightclick-event-excel.md) | 
[Calculate](chart-calculate-event-excel.md) | 
[Deactivate](chart-deactivate-event-excel.md) | 
[MouseDown](chart-mousedown-event-excel.md) | 
[MouseMove](chart-mousemove-event-excel.md) | 
[MouseUp](chart-mouseup-event-excel.md) | 
[Resize](chart-resize-event-excel.md) | 
[Select](chart-select-event-excel.md) | 
[SeriesChange](chart-serieschange-event-excel.md)

 **Note**  To write event procedures for an embedded chart, you must create a new object using the  **WithEvents** keyword in a class module. For more information, see [Using Events with Embedded Charts](using-events-with-embedded-charts.md).

This example changes a point's border color when the user changes the point value.



```vb
Private Sub Chart_SeriesChange(ByVal SeriesIndex As Long, _ 
        ByVal PointIndex As Long) 
    Set p = ActiveChart.SeriesCollection(SeriesIndex). _ 
        Points(PointIndex) 
    p.Border.ColorIndex = 3 
End Sub
```


