---
title: RulerGuides Object (Publisher)
keywords: vbapb10.chm786431
f1_keywords:
- vbapb10.chm786431
ms.prod: publisher
api_name:
- Publisher.RulerGuides
ms.assetid: c58d3cb2-8cf8-74fa-2bf4-a931dc95a26a
ms.date: 06/08/2017
---


# RulerGuides Object (Publisher)

A collection of  **[RulerGuide](rulerguide-object-publisher.md)** objects that represents a gridline used to align objects on a page.
 


## Example

Use the  **[Add](rulerguides-add-method-publisher.md)** method of the **RulerGuides** collection to add ruler gridlines to the **RulerGuides** collection. This example creates horizontal ruler guides and vertical ruler guides every half inch on the first page of the active publication.
 

 

```
Sub SetRulerGuides() 
 Dim intCount As Integer 
 Dim intPos As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To 16 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeVertical 
 Next intCount 
 intPos = 0 
 For intCount = 1 To 21 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeHorizontal 
 Next intCount 
 End With 
End Sub
```

Use the  **[Count](rulerguides-count-property-publisher.md)** property to return the total number of ruler guides, horizontal and vertical, in the collection. The following example uses the **Count** property to create a loop that deletes each of the ruler guides in the collection.
 

 



```
Sub RemoveAllGuides() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To .Count 
 .Item(1).Delete 
 Next intCount 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](rulerguides-add-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](rulerguides-application-property-publisher.md)|
|[Count](rulerguides-count-property-publisher.md)|
|[Item](rulerguides-item-property-publisher.md)|
|[Parent](rulerguides-parent-property-publisher.md)|

