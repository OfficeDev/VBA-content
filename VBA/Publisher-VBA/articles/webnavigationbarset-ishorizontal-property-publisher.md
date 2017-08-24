---
title: WebNavigationBarSet.IsHorizontal Property (Publisher)
keywords: vbapb10.chm8519686
f1_keywords:
- vbapb10.chm8519686
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.IsHorizontal
ms.assetid: d3bbb0b0-8d06-7d46-1ef7-fef0a3e846b7
ms.date: 06/08/2017
---


# WebNavigationBarSet.IsHorizontal Property (Publisher)

 **True** if the specified **WebNavigationBarSet** has a horizontal orientation. Read-only **Boolean**.


## Syntax

 _expression_. **IsHorizontal**

 _expression_A variable that represents an  **WebNavigationBarSet** object.


### Return Value

Boolean


## Remarks

This property is used to determine the orientation of the navigation bar set prior to setting certain properties that assume a horizontal orientation such as the  **HorizontalAlignment** property.


## Example

This example adds the first navigation bar in the  **WebNavigationBarSets** collection of the active document to each page, and then sets the button style to **small**. A test is performed to determine whether the navigation bar set is horizontal or not. If it is not, the  **ChangeOrientation** method is called and the orientation is set to **horizontal**. After the navigation bar is oriented horizontally, the horizontal button count is set to  **3** and the horizontal alignment of the buttons is set to **left**.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleSmall 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignLeft 
End With
```


