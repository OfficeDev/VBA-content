---
title: WebListBox.ListBoxItems Property (Publisher)
keywords: vbapb10.chm4063235
f1_keywords:
- vbapb10.chm4063235
ms.prod: publisher
api_name:
- Publisher.WebListBox.ListBoxItems
ms.assetid: 642a4592-35af-99fa-ee96-6bd8517c618f
ms.date: 06/08/2017
---


# WebListBox.ListBoxItems Property (Publisher)

Returns a  **[WebListBoxItems](weblistboxitems-object-publisher.md)** object that represents the items in a Web list box control.


## Syntax

 _expression_. **ListBoxItems**

 _expression_A variable that represents a  **WebListBox** object.


### Return Value

WebListBoxItems


## Example

This example creates a new Web list box control and adds five new list items to it.


```vb
Sub NewListBoxItems() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlListBox, Left:=100, _ 
 Top:=100, Width:=150, Height:=100).WebListBox 
 .MultiSelect = msoTrue 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Yellow" 
 .AddItem Item:="Red" 
 .AddItem Item:="Blue" 
 .AddItem Item:="Green" 
 .AddItem Item:="Black" 
 End With 
 End With 
End Sub
```


