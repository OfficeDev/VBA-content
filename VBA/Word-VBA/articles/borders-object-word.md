---
title: Borders Object (Word)
ms.prod: word
ms.assetid: 6dd1d4cc-2dcf-22c7-a299-4721a5543ba3
ms.date: 06/08/2017
---


# Borders Object (Word)

A collection of  **[Border](border-object-word.md)** objects that represent the borders of an object.


## Remarks

Use the  **Borders** property to return the **Borders** collection. The following example applies the default border around the first paragraph in the active document.


```
ActiveDocument.Paragraphs(1).Borders.Enable = True
```

 **[Border](border-object-word.md)** objects cannot be added to the **Borders** collection. The number of members in the **Borders** collection is finite and varies depending on the type of object. For example, a table has six elements in the **Borders** collection, whereas a paragraph has four.

Use  **Borders** (index), where index identifies the border, to return a single **Border** object. Index can be one of the **[WdBorderType](wdbordertype-enumeration-word.md)** constants. Some of the **WdBorderType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.

Use the  **[LineStyle](border-linestyle-property-word.md)** property to apply a border line to a **Border** object. The following example applies a double-line border below the first paragraph in the active document.




```
With ActiveDocument.Paragraphs(1).Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleDouble 
 .LineWidth = wdLineWidth025pt 
End With
```

The following example applies a single-line border around the first character in the selection.




```
With Selection.Characters(1) 
 .Font.Size = 36 
 .Borders.Enable = True 
End With
```

The following example adds an art border around each page in the first section.




```
For Each aBorder In ActiveDocument.Sections(1).Borders 
 With aBorder 
 .ArtStyle = wdArtSeattle 
 .ArtWidth = 20 
 End With 
Next aBorder
```


## Methods



|**Name**|
|:-----|
|[ApplyPageBordersToAllSections](borders-applypageborderstoallsections-method-word.md)|
|[Item](borders-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[AlwaysInFront](borders-alwaysinfront-property-word.md)|
|[Application](borders-application-property-word.md)|
|[Count](borders-count-property-word.md)|
|[Creator](borders-creator-property-word.md)|
|[DistanceFrom](borders-distancefrom-property-word.md)|
|[DistanceFromBottom](borders-distancefrombottom-property-word.md)|
|[DistanceFromLeft](borders-distancefromleft-property-word.md)|
|[DistanceFromRight](borders-distancefromright-property-word.md)|
|[DistanceFromTop](borders-distancefromtop-property-word.md)|
|[Enable](borders-enable-property-word.md)|
|[EnableFirstPageInSection](borders-enablefirstpageinsection-property-word.md)|
|[EnableOtherPagesInSection](borders-enableotherpagesinsection-property-word.md)|
|[HasHorizontal](borders-hashorizontal-property-word.md)|
|[HasVertical](borders-hasvertical-property-word.md)|
|[InsideColor](borders-insidecolor-property-word.md)|
|[InsideColorIndex](borders-insidecolorindex-property-word.md)|
|[InsideLineStyle](borders-insidelinestyle-property-word.md)|
|[InsideLineWidth](borders-insidelinewidth-property-word.md)|
|[JoinBorders](borders-joinborders-property-word.md)|
|[OutsideColor](borders-outsidecolor-property-word.md)|
|[OutsideColorIndex](borders-outsidecolorindex-property-word.md)|
|[OutsideLineStyle](borders-outsidelinestyle-property-word.md)|
|[OutsideLineWidth](borders-outsidelinewidth-property-word.md)|
|[Parent](borders-parent-property-word.md)|
|[Shadow](borders-shadow-property-word.md)|
|[SurroundFooter](borders-surroundfooter-property-word.md)|
|[SurroundHeader](borders-surroundheader-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
