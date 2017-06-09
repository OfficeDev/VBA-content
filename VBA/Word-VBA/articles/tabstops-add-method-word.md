---
title: TabStops.Add Method (Word)
keywords: vbawd10.chm156565604
f1_keywords:
- vbawd10.chm156565604
ms.prod: word
api_name:
- Word.TabStops.Add
ms.assetid: d6996a6c-e2e7-692c-3f48-27af213803e1
ms.date: 06/08/2017
---


# TabStops.Add Method (Word)

Returns a  **TabStop** object that represents a custom tab stop added to a document.


## Syntax

 _expression_ . **Add**( **_Position_** , **_Alignment_** , **_Leader_** )

 _expression_ Required. A variable that represents a **[TabStops](tabstops-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Required| **Single**|The position of the tab stop (in points) relative to the left margin.|
| _Alignment_|Optional| **Variant**|The alignment of the tab stop. Can be one of the  **WdTabAlignment** constants.|
| _Leader_|Optional| **Variant**|The type of leader for the tab stop. Can be one of the  **WdTabLeader** constants. If this argument is omitted, **wdTabLeaderSpaces** is used.|

### Return Value

TabStop


## Example

This example adds a tab stop positioned at 2.5 inches (from the left edge of the page) to the selected paragraphs.


```
Selection.Paragraphs.TabStops.Add Position:=InchesToPoints(2.5)
```

This example adds two tab stops to the selected paragraphs. The first tab stop is a left aligned, has a dotted leader, and is positioned at 1 inch (72 points) from the left edge of the page. The second tab stop is centered and is positioned at 2 inches from the left edge.




```vb
With Selection.Paragraphs.TabStops 
 .Add Position:=InchesToPoints(1), _ 
 Leader:=wdTabLeaderDots, _ 
 Alignment:=wdAlignTabLeft 
 .Add Position:=InchesToPoints(2), _ 
 Alignment:=wdAlignTabCenter 
End With
```


## See also


#### Concepts


[TabStops Collection Object](tabstops-object-word.md)

