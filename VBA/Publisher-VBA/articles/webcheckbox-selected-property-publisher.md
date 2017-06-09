---
title: WebCheckBox.Selected Property (Publisher)
keywords: vbapb10.chm4325380
f1_keywords:
- vbapb10.chm4325380
ms.prod: publisher
api_name:
- Publisher.WebCheckBox.Selected
ms.assetid: ad34871d-474d-70ad-6245-ee5a017839c1
ms.date: 06/08/2017
---


# WebCheckBox.Selected Property (Publisher)

Specifies whether a Web check box or option button is selected. Read/write.


## Syntax

 _expression_. **Selected**

 _expression_A variable that represents a  **WebCheckBox** object.


## Remarks

The  **Selected** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example adds a new Web check box to the first page of the active publication and then selects it.


```vb
Sub AddNewWebCheckBox() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=100, Width:=100, Height:=12) 
 .WebCheckBox.Selected = msoTrue 
 End With 
End Sub
```


