---
title: Application.SnapToObjects Property (Publisher)
keywords: vbapb10.chm131111
f1_keywords:
- vbapb10.chm131111
ms.prod: publisher
api_name:
- Publisher.Application.SnapToObjects
ms.assetid: 84fcb808-bf3b-49f7-666e-915ac6b04a96
ms.date: 06/08/2017
---


# Application.SnapToObjects Property (Publisher)

 **True** for Microsoft Publisher to use objects on a page to line up other objects. Read/write **Boolean**.


## Syntax

 _expression_. **SnapToObjects**

 _expression_A variable that represents a  **Application** object.


### Return Value

Boolean


## Example

This example adds horizontal and vertical ruler guides every half inch on the first page and sets the options to align objects on the page to the guides.


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

