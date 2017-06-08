---
title: TabStop Object (Publisher)
keywords: vbapb10.chm5701631
f1_keywords:
- vbapb10.chm5701631
ms.prod: publisher
api_name:
- Publisher.TabStop
ms.assetid: 74e71d75-503f-ef57-ddeb-24a788402df2
ms.date: 06/08/2017
---


# TabStop Object (Publisher)

Represents a single tab stop. The  **TabStop** object is a member of the **[TabStops](tabstops-object-publisher.md)** collection. The **TabStops** collection represents all the custom and default tab stops in a paragraph or group of paragraphs.
 


## Remarks

Set the  **[DefaultTabStop](document-defaulttabstop-property-publisher.md)** property to adjust the spacing of default tab stops.
 

 

## Example

Use  **[Tabs](tabstops-add-method-publisher.md)** (index), where index is the location of the tab stop (in points) or the index number, to return a single **TabStop** object. Tab stops are indexed numerically from left to right along the ruler. The following example removes the first custom tab stop from the selected paragraphs.
 

 

```
Sub ClearTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear 
End Sub
```

The following example adds a right-aligned tab stop positioned at 2 inches to the selected paragraphs.
 

 



```
Sub ChangeTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(2) _ 
 .Alignment = pbTabAlignmentTrailing 
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


## Methods



|**Name**|
|:-----|
|[Clear](tabstop-clear-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](tabstop-alignment-property-publisher.md)|
|[Application](tabstop-application-property-publisher.md)|
|[Leader](tabstop-leader-property-publisher.md)|
|[Parent](tabstop-parent-property-publisher.md)|
|[Position](tabstop-position-property-publisher.md)|

