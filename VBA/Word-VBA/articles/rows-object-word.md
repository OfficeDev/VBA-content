---
title: Rows Object (Word)
ms.prod: word
ms.assetid: cd83d0ef-f743-1886-54de-497017c5f542
ms.date: 06/08/2017
---


# Rows Object (Word)

A collection of  **[Row](row-object-word.md)** objects that represent the table rows in the specified selection, range, or table.


## Remarks

Use the  **Rows** property to return the **Rows** collection. The following example centers rows in the first table in the active document between the left and right margins.


```
ActiveDocument.Tables(1).Rows.Alignment = wdAlignRowCenter
```

Use the  **Add** method to add a row to a table. The following example inserts a row before the first row in the selection.




```
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows.Add BeforeRow:=Selection.Rows(1) 
End If
```

Use  **Rows** (Index), where Index is the index number, to return a single **Row** object. The index number represents the position of the row in the selection, range, or table. The following example deletes the first row in the first table in the active document.




```
ActiveDocument.Tables(1).Rows(1).Delete
```


## Methods



|**Name**|
|:-----|
|[Add](rows-add-method-word.md)|
|[ConvertToText](rows-converttotext-method-word.md)|
|[Delete](rows-delete-method-word.md)|
|[DistributeHeight](rows-distributeheight-method-word.md)|
|[Item](rows-item-method-word.md)|
|[Select](rows-select-method-word.md)|
|[SetHeight](rows-setheight-method-word.md)|
|[SetLeftIndent](rows-setleftindent-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](rows-alignment-property-word.md)|
|[AllowBreakAcrossPages](rows-allowbreakacrosspages-property-word.md)|
|[AllowOverlap](rows-allowoverlap-property-word.md)|
|[Application](rows-application-property-word.md)|
|[Borders](rows-borders-property-word.md)|
|[Count](rows-count-property-word.md)|
|[Creator](rows-creator-property-word.md)|
|[DistanceBottom](rows-distancebottom-property-word.md)|
|[DistanceLeft](rows-distanceleft-property-word.md)|
|[DistanceRight](rows-distanceright-property-word.md)|
|[DistanceTop](rows-distancetop-property-word.md)|
|[First](rows-first-property-word.md)|
|[HeadingFormat](rows-headingformat-property-word.md)|
|[Height](rows-height-property-word.md)|
|[HeightRule](rows-heightrule-property-word.md)|
|[HorizontalPosition](rows-horizontalposition-property-word.md)|
|[Last](rows-last-property-word.md)|
|[LeftIndent](rows-leftindent-property-word.md)|
|[NestingLevel](rows-nestinglevel-property-word.md)|
|[Parent](rows-parent-property-word.md)|
|[RelativeHorizontalPosition](rows-relativehorizontalposition-property-word.md)|
|[RelativeVerticalPosition](rows-relativeverticalposition-property-word.md)|
|[Shading](rows-shading-property-word.md)|
|[SpaceBetweenColumns](rows-spacebetweencolumns-property-word.md)|
|[TableDirection](rows-tabledirection-property-word.md)|
|[VerticalPosition](rows-verticalposition-property-word.md)|
|[WrapAroundText](rows-wraparoundtext-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
