---
title: TableStyle Object (Word)
keywords: vbawd10.chm3735
f1_keywords:
- vbawd10.chm3735
ms.prod: word
api_name:
- Word.TableStyle
ms.assetid: 4f1f4489-0ef7-dff0-8f2a-77f87937f3ad
ms.date: 06/08/2017
---


# TableStyle Object (Word)

Represents a single style that can be applied to a table.


## Remarks

Use the  **Table** property of the **Styles** object to return a **TableStyle** object. Use the **Borders** property to apply borders to an entire table. Use the **Condition** method to apply borders or shading only to specified sections of a table. This example creates a new table style and formats the table with a surrounding border. Special borders and shading are applied to the first and last rows and the last column.


```
Sub NewTableStyle() 
 Dim styTable As Style 
 
 Set styTable = ActiveDocument.Styles.Add( _ 
 Name:="TableStyle 1", Type:=wdStyleTypeTable) 
 
 With styTable.Table 
 
 'Apply borders around table 
 .Borders(wdBorderTop).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle 
 .Borders(wdBorderRight).LineStyle = wdLineStyleSingle 
 
 'Apply a double border to the heading row 
 .Condition(wdFirstRow).Borders(wdBorderBottom) _ 
 .LineStyle = wdLineStyleDouble 
 
 'Apply a double border to the last column 
 .Condition(wdLastColumn).Borders(wdBorderLeft) _ 
 .LineStyle = wdLineStyleDouble 
 
 'Apply shading to last row 
 .Condition(wdLastRow).Shading _ 
 .BackgroundPatternColor = wdColorGray125 
 
 End With 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Condition](tablestyle-condition-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](tablestyle-alignment-property-word.md)|
|[AllowBreakAcrossPage](tablestyle-allowbreakacrosspage-property-word.md)|
|[AllowPageBreaks](tablestyle-allowpagebreaks-property-word.md)|
|[Application](tablestyle-application-property-word.md)|
|[Borders](tablestyle-borders-property-word.md)|
|[BottomPadding](tablestyle-bottompadding-property-word.md)|
|[ColumnStripe](tablestyle-columnstripe-property-word.md)|
|[Creator](tablestyle-creator-property-word.md)|
|[LeftIndent](tablestyle-leftindent-property-word.md)|
|[LeftPadding](tablestyle-leftpadding-property-word.md)|
|[Parent](tablestyle-parent-property-word.md)|
|[RightPadding](tablestyle-rightpadding-property-word.md)|
|[RowStripe](tablestyle-rowstripe-property-word.md)|
|[Shading](tablestyle-shading-property-word.md)|
|[Spacing](tablestyle-spacing-property-word.md)|
|[TableDirection](tablestyle-tabledirection-property-word.md)|
|[TopPadding](tablestyle-toppadding-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
