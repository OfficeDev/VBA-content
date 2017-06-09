---
title: WebNavigationBarSet.ShowSelected Property (Publisher)
keywords: vbapb10.chm8519696
f1_keywords:
- vbapb10.chm8519696
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.ShowSelected
ms.assetid: c8229f03-a043-a280-84f9-f75a430c3903
ms.date: 06/08/2017
---


# WebNavigationBarSet.ShowSelected Property (Publisher)

 **True** if the selected button is highlighted for the specified **WebNavigationBarSet** object. Read/write **Boolean**.


## Syntax

 _expression_. **ShowSelected**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Return Value

Boolean


## Example

The following example adds a new Web navigation bar to the active document, adds it to every page, and then sets the  **ShowSelected** property to **False** so that the selected button will not be highlighted in the navigation bar.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newNavBar") 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleSmall 
 .ShowSelected = False 
End With
```


