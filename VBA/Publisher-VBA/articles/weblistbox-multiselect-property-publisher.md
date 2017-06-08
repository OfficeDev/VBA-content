---
title: WebListBox.MultiSelect Property (Publisher)
keywords: vbapb10.chm4063236
f1_keywords:
- vbapb10.chm4063236
ms.prod: publisher
api_name:
- Publisher.WebListBox.MultiSelect
ms.assetid: cc81682f-5212-0912-d979-16567c2dc42b
ms.date: 06/08/2017
---


# WebListBox.MultiSelect Property (Publisher)

Specifies whether a user may select more than one item in a Web list box control. Read/write.


## Syntax

 _expression_. **MultiSelect**

 _expression_A variable that represents a  **WebListBox** object.


### Return Value

MsoTriState


## Remarks

The  **MultiSelect** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Indicates a user may only select one item in a Web list box control.|
| **msoTrue**| Indicates a user may select more than one item in a Web list box control.|

## Example

This example add a Web list box control to the active publication, add items to it, and specifies that a user may select more than one item.


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


