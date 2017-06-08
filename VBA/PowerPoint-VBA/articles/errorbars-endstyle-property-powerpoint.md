---
title: ErrorBars.EndStyle Property (PowerPoint)
keywords: vbapp10.chm66660
f1_keywords:
- vbapp10.chm66660
ms.prod: powerpoint
api_name:
- PowerPoint.ErrorBars.EndStyle
ms.assetid: 2d6dca80-0281-3d56-fdc9-dc863bf7ad38
ms.date: 06/08/2017
---


# ErrorBars.EndStyle Property (PowerPoint)

Returns or sets the end style for the error bars. Read/write  **Long**.


## Syntax

 _expression_. **EndStyle**

 _expression_ A variable that represents an **[ErrorBars](errorbars-object-powerpoint.md)** object.


## Remarks

The value of this property can be one of the following  **[XlEndStyleCap](xlendstylecap-enumeration-powerpoint.md)** constants:


-  **xlCap**
    
-  **xlNoCap**
    



## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the end style for the error bars for series one of the first chart in the active document. You should run the example on a 2-D line chart that has Y error bars for the first series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).ErrorBars.EndStyle = xlCap

    End If

End With
```


## See also


#### Concepts


[ErrorBars Object](errorbars-object-powerpoint.md)


