---
title: Font Object
keywords: vbagr10.chm131085
f1_keywords:
- vbagr10.chm131085
ms.prod: excel
api_name:
- Excel.Font
ms.assetid: 0510e805-48fd-7148-edee-d65dc59f34b4
ms.date: 06/08/2017
---


# Font Object

Contains the font attributes (font name, font size, color, and so on) for the specified object.


## Using the Font Object

Use the  **Font** property to return the **Font** object. The following example sets the title text for the value axis, sets the font to 10-point Bookman, and formats the word "millions" as italic.


```vb
With myChart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 .Characters(10, 8).Font.Italic = True
 End With 
End With
```


