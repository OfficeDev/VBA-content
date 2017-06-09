---
title: WebNavigationBarSet.AutoUpdate Property (Publisher)
keywords: vbapb10.chm8519689
f1_keywords:
- vbapb10.chm8519689
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.AutoUpdate
ms.assetid: b9ce8dde-c09f-6fe9-6935-cb4903a17b85
ms.date: 06/08/2017
---


# WebNavigationBarSet.AutoUpdate Property (Publisher)

 **True** if all pages will be added to the specified Web navigation bar set and that adding new pages will update the navigation bar with a corresponding item. Pages must have the **AddHyperlinkToWebNavbar** set to **True** or **WebPageOptions.IncludePageOnNewWebNavigationBars** property set to **True** to be added or updated within the specified **WebNavigationBarSet**. Read/write  **Boolean**.


## Syntax

 _expression_. **AutoUpdate**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Return Value

Boolean


## Remarks

This property determines whether or not the existing pages in the publication will be added to the navigation bar and if added pages will also be updated. These pages must be marked with the  **AddHyperlinkToWebNavbar** set to **True** or **WebPageOptions.IncludePageOnNewWebNavigationBars** property set to **True** to be added or updated within the specified **WebNavigationBarSet**. Changing this setting does not change the number of items in the bar, it just determines whether or not new pages will be added. By setting this value to  **False** it is possible to design specific navigation bars for specific content pages in a Web site that do not contain all of the available hyperlinks in the publication.

The default value is  **True**. 


## Example

The following example adds a new Web navigation bar set to the active document with text style buttons and auto update set to  **False** so that page links will not be added or new pages automatically updated in the navigation bar, then the Web navigation bar is added to the first page of the publication.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newBar") 
With objWebNav 
 .AutoUpdate = False 
 .ButtonStyle = pbnbButtonStyleText 
End With 
ActiveDocument.Pages(1).Shapes.AddWebNavigationBar _ 
 Name:="newBar", Left:=10, Top:=10 

```


