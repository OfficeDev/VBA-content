---
title: View.ActivePage Property (Publisher)
keywords: vbapb10.chm327683
f1_keywords:
- vbapb10.chm327683
ms.prod: publisher
api_name:
- Publisher.View.ActivePage
ms.assetid: 29289fb2-6692-4cb5-a9e2-b2edb9e9cd7e
ms.date: 06/08/2017
---


# View.ActivePage Property (Publisher)

Returns a  **[Page](page-object-publisher.md)** object that represents the page currently displayed in the Microsoft Publisher window.


## Syntax

 _expression_. **ActivePage**

 _expression_A variable that represents a  **View** object.


### Return Value

Page


## Example

This example saves the active page as a JPEG picture. (Note that PathToFile must be replaced with a valid file path for this example to work.)


```vb
Sub SavePageAsPicture() 
 ActiveView.ActivePage.SaveAsPicture _ 
 FileName:="PathToFile" 
End Sub
```

This example adds a horizontal ruler guide and a vertical ruler guide to the active page that intersect at the center point of the page.




```vb
Sub SetRulerGuidesOnActivePage() 
 Dim intHeight As Integer 
 Dim intWidth As Integer 
 
 With ActiveView.ActivePage 
 intHeight = .Height / 2 
 intWidth = .Width / 2 
 With .RulerGuides 
 .Add Position:=intHeight, Type:=pbRulerGuideTypeHorizontal 
 .Add Position:=intWidth, Type:=pbRulerGuideTypeVertical 
 End With 
 End With 
End Sub
```


