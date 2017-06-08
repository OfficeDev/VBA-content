---
title: TabStops Object (Word)
keywords: vbawd10.chm2389
f1_keywords:
- vbawd10.chm2389
ms.prod: word
ms.assetid: 2d3bcac4-db8c-05fe-1cc1-5d90774f84fb
ms.date: 06/08/2017
---


# TabStops Object (Word)

A collection of  **[TabStop](tabstop-object-word.md)** objects that represent the custom and default tabs for a paragraph or group of paragraphs.


## Remarks

Use the  **TabStops** property to return the **TabStops** collection. The following example clears all the custom tab stops from the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).TabStops.ClearAll
```

The following example adds a tab stop positioned at 2.5 inches to the selected paragraphs and then displays the position of each item in the  **TabStops** collection.




```vb
Selection.Paragraphs.TabStops.Add Position:=InchesToPoints(2.5) 
For Each aTab In Selection.Paragraphs.TabStops 
 MsgBox "Position = " _ 
 &; PointsToInches(aTab.Position) &; " inches" 
Next aTab
```

Use the  **Add** method to add a tab stop. The following example adds two tab stops to the selected paragraphs. The first tab stop is a left-aligned tab with a dotted tab leader positioned at 1 inch (72 points). The second tab stop is centered and is positioned at 2 inches.




```vb
With Selection.Paragraphs.TabStops 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=wdTabLeaderDots, Alignment:=wdAlignTabLeft 
 .Add Position:=InchesToPoints(2), Alignment:=wdAlignTabCenter 
End With
```

You can also add a tab stop by specifying a location with the  **TabStops** property. The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.




```
Selection.Paragraphs.TabStops(InchesToPoints(2)) _ 
 .Alignment = wdAlignTabRight
```

Use  **TabStops** (Index), where Index is the location of the tab stop (in points) or the index number, to return a single **TabStop** object. Tab stops are indexed numerically from left to right along the ruler. The following example removes the first custom tab stop from the first paragraph in the active document.




```vb
ActiveDocument.Paragraphs(1).TabStops(1).Clear
```

The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.




```
Selection.Paragraphs.TabStops(InchesToPoints(2)) _ 
 .Alignment = wdAlignTabRight
```

When working with the  **Paragraphs** collection (or a range with several paragraphs), you must modify each paragraph in the collection individually if the tab stops aren't identical in all the paragraphs. The following example removes the tab positioned at 1 inch from every paragraph in the active document.




```vb
For Each para In ActiveDocument.Content.Paragraphs 
 para.TabStops(InchesToPoints(1)).Clear 
Next para
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


