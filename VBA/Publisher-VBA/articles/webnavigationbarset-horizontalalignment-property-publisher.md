---
title: WebNavigationBarSet.HorizontalAlignment Property (Publisher)
keywords: vbapb10.chm8519688
f1_keywords:
- vbapb10.chm8519688
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.HorizontalAlignment
ms.assetid: 7d615a5a-793c-fd78-3dca-a268740b67aa
ms.date: 06/08/2017
---


# WebNavigationBarSet.HorizontalAlignment Property (Publisher)

Sets or returns a  **PbWizardNavBarAlignment** constant that represents the horizontal alignment of the buttons in a Web navigation bar set. Read/write.


## Syntax

 _expression_. **HorizontalAlignment**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Return Value

PbWizardNavBarAlignment


## Remarks

This property is used to set the way that buttons are displayed in a horizontally oriented Web navigation bar set. For example, a  **WebNavigationBarSet** object containing 5 links with the **HorizontalButtonCount** property set to 3 and the **HorizontalAlignment** property set to **pbnbAlignRight** will align the buttons in a grid of 3 columns and 1 row. The first 3 buttons will be in the first row and the remaining 2 buttons will be in the rightmost columns of the second row.

Returns "Access denied" if  **IsHorizontal** = **False** for the specified **WebNavigationBarSet** object. Use the **ChangeOrientation** method to set the orientation of the Web navigation bar set to horizontal first before setting the **HorizontalAlignment** property.

The  **HorizontalAlignment** property value can be set to any of the **[PbWizardNavBarAlignment](pbwizardnavbaralignment-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

The following example returns the first Web navigation bar set from the active document, changes the orientation to horizontal if necessary, sets the  **HorizontalButtonCount** property to 3, and then sets the **HorizontalAlignment** property to **pbnbAlignRight**.


```vb
With ActiveDocument.WebNavigationBarSets(1) 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignRight 
End With
```


