---
title: Application.SnapToGuides Property (Publisher)
keywords: vbapb10.chm131110
f1_keywords:
- vbapb10.chm131110
ms.prod: publisher
api_name:
- Publisher.Application.SnapToGuides
ms.assetid: 09894c02-3193-cd14-ff55-45920e461af9
ms.date: 06/08/2017
---


# Application.SnapToGuides Property (Publisher)

 **True** for Microsoft Publisher to use the guides to align objects on a page in a publication. Read/write **Boolean**.


## Syntax

 _expression_. **SnapToGuides**

 _expression_A variable that represents a  **Application** object.


### Return Value

Boolean


## Example

This example adds horizontal and vertical ruler guides every half inch on the first page and then sets the options to align objects on the page to the guides.


```vb
Sub SetSnapOptions() 
 Dim intCount As Integer 
 Dim intPos As Integer 
 With ActiveDocument.Pages(1).RulerGuides 
 For intCount = 1 To 16 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeVertical 
 Next 
 intPos = 0 
 For intCount = 1 To 21 
 intPos = intPos + 36 
 .Add Position:=intPos, Type:=pbRulerGuideTypeHorizontal 
 Next 
 End With 
 With Application 
 .SnapToGuides = True 
 .SnapToObjects = True 
 End With 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

