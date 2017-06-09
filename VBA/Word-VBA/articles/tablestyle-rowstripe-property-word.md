---
title: TableStyle.RowStripe Property (Word)
keywords: vbawd10.chm244776977
f1_keywords:
- vbawd10.chm244776977
ms.prod: word
api_name:
- Word.TableStyle.RowStripe
ms.assetid: 4ffb0d84-d2a6-bf22-39e0-bd24d9a41ef0
ms.date: 06/08/2017
---


# TableStyle.RowStripe Property (Word)

Returns or sets a  **Long** that represents the number of rows to include in the banding when a style specifies odd- or even-row banding. Read/write.


## Syntax

 _expression_ . **RowStripe**

 _expression_ A variable that represents a **[TableStyle](tablestyle-object-word.md)** object.


## Remarks

Use the  **[Condition](tablestyle-condition-method-word.md)** method to set odd- or even-column banding for a table style.


## Example

This example creates and formats a new table style then applies the new style to a new table. The resulting style causes three columns every third column and two rows every second row to have 20% shading.


```vb
Sub NewTableStyle() 
 Dim styTable As Style 
 
 With ActiveDocument 
 Set styTable = .Styles.Add(Name:="TableStyle 1", _ 
 Type:=wdStyleTypeTable) 
 
 With .Styles("TableStyle 1").Table 
 .Condition(wdEvenColumnBanding).Shading _ 
 .Texture = wdTexture20Percent 
 .ColumnStripe = 3 
 .Condition(wdEvenRowBanding).Shading _ 
 .Texture = wdTexture20Percent 
 .RowStripe = 2 
 End With 
 
 With .Tables.Add(Range:=Selection.Range, NumRows:=15, _ 
 NumColumns:=15) 
 .Style = ActiveDocument.Styles("TableStyle 1") 
 End With 
 End With 
 
End Sub
```


## See also


#### Concepts


[TableStyle Object](tablestyle-object-word.md)

