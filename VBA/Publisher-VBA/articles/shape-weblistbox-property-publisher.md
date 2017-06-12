---
title: Shape.WebListBox Property (Publisher)
keywords: vbapb10.chm2228341
f1_keywords:
- vbapb10.chm2228341
ms.prod: publisher
api_name:
- Publisher.Shape.WebListBox
ms.assetid: c100dfc7-6fbd-db48-4de9-4a9a49739a8f
ms.date: 06/08/2017
---


# Shape.WebListBox Property (Publisher)

Returns the  **[WebListBox](weblistbox-object-publisher.md)** object associated with the specified shape.


## Syntax

 _expression_. **WebListBox**

 _expression_A variable that represents a  **Shape** object.


### Return Value

WebListBox


## Example

This example creates a new Web list box and adds several items to it. Note that when initially created, a Web list box control contains three default items. This example includes a loop that deletes the default list box items before adding new items.


```vb
Dim shpNew As Shape 
Dim wlbTemp As WebListBox 
Dim intCount As Integer 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlListBox, Left:=100, _ 
 Top:=150, Width:=300, Height:=72) 
 
Set wlbTemp = shpNew.Web ListBox 
 
With wlbTemp 
 .MultiSelect = msoFalse 
 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next intCount 
 
 .AddItem Item:="Green" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Red" 
 .AddItem Item:="Black" 
 End With 
End With
```


