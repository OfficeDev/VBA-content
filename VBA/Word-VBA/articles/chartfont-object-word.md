---
title: ChartFont Object (Word)
keywords: vbawd10.chm3905
f1_keywords:
- vbawd10.chm3905
ms.prod: word
api_name:
- Word.ChartFont
ms.assetid: 2ca7fb97-fa22-dec1-6978-8ebb6d8aad7c
ms.date: 06/08/2017
---


# ChartFont Object (Word)

Contains the font attributes (font name, font size, color, and so on) for an object chart.


## Remarks

If you do not want to format all the text in an  **[AxisTitle](axistitle-object-word.md)** , **[ChartTitle](charttitle-object-word.md)** , **[DataLabel](datalabel-object-word.md)** , or **[DisplayUnitLabel](displayunitlabel-object-word.md)** object the same way, use the **Characters** property of that object to first return a subset of the text as a **[ChartCharacters](chartcharacters-object-word.md)** object. Then use the **[Font](chartcharacters-font-property-word.md)** property of the **ChartCharacters** object to return a **ChartFont** object you can use to format the subset of text, as needed.


## Example

The following example formats the title of the first chart as bold. Use the  **Font** property to return the **ChartFont** object.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .AxisTitle.Font.Bold = True 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

