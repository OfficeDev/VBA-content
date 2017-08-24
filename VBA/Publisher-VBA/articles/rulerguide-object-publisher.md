---
title: RulerGuide Object (Publisher)
keywords: vbapb10.chm720895
f1_keywords:
- vbapb10.chm720895
ms.prod: publisher
api_name:
- Publisher.RulerGuide
ms.assetid: 6400c368-02e9-169c-c675-9416cd361384
ms.date: 06/08/2017
---


# RulerGuide Object (Publisher)

Represents a gridline used to align objects on a page. The  **RulerGuide** object is a member of the **[RulerGuides](rulerguides-object-publisher.md)** collection.
 


## Example

Use the  **[Add](rulerguides-add-method-publisher.md)** method of the **RulerGuides** collection to create a new ruler gridline. Use the **[Item](rulerguides-item-property-publisher.md)** property to reference a ruler guide. Use the **[Position](rulerguide-position-property-publisher.md)** property to change the position of a gridline, and use the **[Delete](rulerguide-delete-method-publisher.md)** method to remove a gridline. This example creates a new ruler guide, moves it, and then deletes it.
 

 

```
Sub AddChangeDeleteGuide() 
 Dim rgLine As RulerGuide 
 With ActiveDocument.Pages(1).RulerGuides 
 .Add Position:=InchesToPoints(1), _ 
 Type:=pbRulerGuideTypeVertical 
 
 MsgBox "The ruler guide position is at one inch." 
 
 .Item(1).Position = InchesToPoints(3) 
 MsgBox "The ruler guide is now at three inches." 
 
 .Item(1).Delete 
 MsgBox "The ruler guide has been deleted." 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](rulerguide-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](rulerguide-application-property-publisher.md)|
|[Parent](rulerguide-parent-property-publisher.md)|
|[Position](rulerguide-position-property-publisher.md)|
|[Type](rulerguide-type-property-publisher.md)|

