---
title: TabStops Object (Publisher)
keywords: vbapb10.chm5636095
f1_keywords:
- vbapb10.chm5636095
ms.prod: publisher
api_name:
- Publisher.TabStops
ms.assetid: fbaa194c-754a-3437-c3d5-fa70c951ca4f
ms.date: 06/08/2017
---


# TabStops Object (Publisher)

A collection of  **[TabStop](tabstop-object-publisher.md)** objects that represent the custom and default tabs for a paragraph or group of paragraphs.
 


## Example

Use the  **[Tabs](paragraphformat-tabs-property-publisher.md)** property to return the **TabStops** collection. The following example clears all the custom tab stops from the first paragraph in the active publication.
 

 

```
Sub ClearAllTabStops() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.Tabs.ClearAll 
End Sub
```

The following example adds a tab stop positioned at 2.5 inches to the selected paragraphs and then displays the position of each item in the  **TabStops** collection.
 

 



```
Sub Tabs() 
 Dim intTab As Integer 
 Selection.TextRange.ParagraphFormat.Tabs _ 
 .Add Position:=InchesToPoints(2.5), _ 
 Alignment:=pbTabAlignmentLeading, Leader:=pbTabLeaderNone 
 With Selection.TextRange.ParagraphFormat 
 For intTab = 1 To .Tabs.Count 
 MsgBox "Position = " &amp; PointsToInches _ 
 (.Tabs(intTab).Position) &amp; " inches" 
 intTab = intTab + 1 
 Next intTab 
 End With 
End Sub
```

Use the  **[Add](tabstops-add-method-publisher.md)** method to add a tab stop. The following example adds two tab stops to the selected paragraphs. The first tab stop is a left-aligned tab with a dotted tab leader positioned at 1 inch (72 points). The second tab stop is centered and is positioned at 2 inches.
 

 



```
Sub AddNewTabs() 
 With Selection.TextRange.ParagraphFormat.Tabs 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=pbTabLeaderDot, Alignment:=pbTabAlignmentLeading 
 .Add Position:=InchesToPoints(2), _ 
 Leader:=pbTabLeaderNone, Alignment:=pbTabAlignmentCenter 
 End With 
End Sub
```

Use  **[Tabs](paragraphformat-tabs-property-publisher.md)** (index), where index is the location of the tab stop (in points) or the index number, to return a single **TabStop** object. Tab stops are indexed numerically from left to right along the ruler. The following example removes the first custom tab stop from the first paragraph in the active publication.
 

 



```
Sub ClearTabStop() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.Tabs(1).Clear 
End Sub
```

The following example changes the second tab in the selection to a right-aligned tab stop.
 

 



```
Sub ChangeTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(2) _ 
 .Alignment = pbTabAlignmentTrailing 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](tabstops-add-method-publisher.md)|
|[ClearAll](tabstops-clearall-method-publisher.md)|
|[Item](tabstops-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](tabstops-application-property-publisher.md)|
|[Count](tabstops-count-property-publisher.md)|
|[Parent](tabstops-parent-property-publisher.md)|

