---
title: WebListBoxItems Object (Publisher)
keywords: vbapb10.chm4194303
f1_keywords:
- vbapb10.chm4194303
ms.prod: publisher
api_name:
- Publisher.WebListBoxItems
ms.assetid: 6d1b6755-426b-b518-c95c-7b30f9acceba
ms.date: 06/08/2017
---


# WebListBoxItems Object (Publisher)

Represents the items in a Web list box control.
 


## Example

Use the  **[ListBoxItems](weblistbox-listboxitems-property-publisher.md)** property to access the items in a Web list box. Use the **[AddItem](weblistboxitems-additem-method-publisher.md)** method of the **WebListBoxItems** collection to add items to a Web list box. This example creates a new Web list box and adds several items to it. Note that when initially created, a Web list box control contains three default items. This example includes a routine that deletes the default list box items before adding new items.
 

 

```
Sub CreateWebListBox() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlListBox, Left:=100, _ 
 Top:=150, Width:=300, Height:=72).WebListBox 
 .MultiSelect = msoFalse 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Green" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Red" 
 .AddItem Item:="Black" 
 End With 
 End With 
 End With 
End Sub
```


