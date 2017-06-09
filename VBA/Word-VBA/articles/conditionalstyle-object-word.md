---
title: ConditionalStyle Object (Word)
keywords: vbawd10.chm1389
f1_keywords:
- vbawd10.chm1389
ms.prod: word
api_name:
- Word.ConditionalStyle
ms.assetid: 2380494e-09e9-8494-a93c-8bbaf621aad1
ms.date: 06/08/2017
---


# ConditionalStyle Object (Word)

Represents special formatting applied to specified areas of a table when the selected table is formatted with a specified table style.


## Remarks

Use the  **[Condition](tablestyle-condition-method-word.md)** method of the **[TableStyle](tablestyle-object-word.md)** object to return a **ConditionalStyle** object. The **Shading** property can be used to apply shading to specified areas of a table. This example selects the first table in the active document and applies shading to alternate rows and columns. This example assumes that there is a table in the active document and that it is formatted using the Table Grid style.


```vb
Sub ApplyConditionalStyle() 
 With ActiveDocument 
 .Tables(1).Select 
 With .Styles("Table Grid").Table 
 .Condition(wdOddColumnBanding).Shading _ 
 .BackgroundPatternColor = wdColorGray10 
 .Condition(wdOddRowBanding).Shading _ 
 .BackgroundPatternColor = wdColorGray10 
 End With 
 End With 
End Sub
```

Use the  **[Borders](tablestyle-borders-property-word.md)** property to apply borders to specified areas of a table. This example selects the first table in the active document and applies borders to the first and last row and first column. This example assumes that there is a table in the active document and that it is formatted using the Table Grid style.




```vb
Sub ApplyTableBorders() 
 With ActiveDocument 
 .Tables(1).Select 
 With .Styles("Table Grid").Table 
 .Condition(wdFirstRow).Borders(wdBorderBottom) _ 
 .LineStyle = wdLineStyleDouble 
 .Condition(wdFirstColumn).Borders(wdBorderRight) _ 
 .LineStyle = wdLineStyleDouble 
 .Condition(wdLastRow).Borders(wdBorderTop) _ 
 .LineStyle = wdLineStyleDouble 
 End With 
 End With 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


