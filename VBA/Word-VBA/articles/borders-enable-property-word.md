---
title: Borders.Enable Property (Word)
keywords: vbawd10.chm154927106
f1_keywords:
- vbawd10.chm154927106
ms.prod: word
api_name:
- Word.Borders.Enable
ms.assetid: 5595b02a-35a3-30ce-0b9b-e6e5867d7258
ms.date: 06/08/2017
---


# Borders.Enable Property (Word)

Returns or sets border formatting for the specified object. Read/write  **Long** .


## Syntax

 _expression_ . **Enable**

 _expression_ A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

The  **Enable** property returns **True** or **wdUndefined** if border formatting is applied to all or part of the specified object. Can be set to **True** , **False** , or a **WdLineStyle** constant.

The  **Enable** property applies to all borders for the specified object. **True** sets the line style to the default line style and sets the line width to the default line width. The default line style and line width can be set using the **DefaultBorderLineWidth** and **DefaultBorderLineStyle** properties.

To remove all the borders from an object, set the  **Enable** property to **False** , as shown in the following example.




```vb
ActiveDocument.Tables(1).Borders.Enable = False
```

To remove or apply a single border, use  **Borders** ( _index_ ), where **index** is a **WdBorderType** constant, to return a single border, and then set the **LineStyle** property. The following example removes the bottom border from `rngTemp`.




```vb
Dim rngTemp 
 
rngTemp.Borders(wdBorderBottom).LineStyle = wdLineStyleNone 

```


## Example

This example removes all borders from the first cell in table one.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 ActiveDocument.Tables(1).Cell(1, 1).Borders.Enable = False 
End If
```

This example applies a dashed border around the first paragraph in the selection.




```
Options.DefaultBorderLineWidth = wdLineWidth025pt 
Selection.Paragraphs(1).Borders.Enable = _ 
 wdLineStyleDashSmallGap
```

This example applies a border around the first character in the selection. If nothing is selected, the border is applied to the first character after the insertion point.




```vb
Selection.Characters(1).Borders.Enable = True
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

