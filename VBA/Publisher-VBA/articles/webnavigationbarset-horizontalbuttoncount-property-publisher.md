---
title: WebNavigationBarSet.HorizontalButtonCount Property (Publisher)
keywords: vbapb10.chm8519687
f1_keywords:
- vbapb10.chm8519687
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.HorizontalButtonCount
ms.assetid: 2f6c5258-16c9-19fd-16c6-ea59c561e9de
ms.date: 06/08/2017
---


# WebNavigationBarSet.HorizontalButtonCount Property (Publisher)

Sets or returns a  **Long** representing the number of buttons in each row of buttons for a Web navigation bar set. Read/write. **Long**.


## Syntax

 _expression_. **HorizontalButtonCount**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Return Value

Long


## Remarks

Returns "Access denied" if  **IsHorizontal** = **False** for the specified **WebNavigationBarSet** object. Use the **ChangeOrientation** method to set the orientation of the Web navigation bar set to **horizontal** first before setting the **HorizontalButtonCount** property.


## Example

The following example returns the first Web navigation bar set from the active document, changes the orientation to  **horizontal** if necessary, sets the **HorizontalButtonCount** property to **3**, and then sets the  **HorizontalAlignment** property to **pbnbAlignLeft**.


```vb
With ActiveDocument.WebNavigationBarSets(1) 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignRight 
End With
```


