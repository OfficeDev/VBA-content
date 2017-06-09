---
title: Page.RulerGuides Property (Publisher)
keywords: vbapb10.chm393225
f1_keywords:
- vbapb10.chm393225
ms.prod: publisher
api_name:
- Publisher.Page.RulerGuides
ms.assetid: 69605642-7722-0721-cb07-d33689eda9ab
ms.date: 06/08/2017
---


# Page.RulerGuides Property (Publisher)

Returns a  **[RulerGuides](rulerguides-object-publisher.md)** collection that represents gridlines used to align objects on a page.


## Syntax

 _expression_. **RulerGuides**

 _expression_A variable that represents a  **Page** object.


### Return Value

RulerGuides


## Example

This example creates horizontal ruler guides and vertical ruler guides every half inch on the first page of the active publication.


```vb
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


