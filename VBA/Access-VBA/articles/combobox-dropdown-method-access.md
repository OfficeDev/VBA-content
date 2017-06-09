---
title: ComboBox.Dropdown Method (Access)
keywords: vbaac10.chm11359
f1_keywords:
- vbaac10.chm11359
ms.prod: access
api_name:
- Access.ComboBox.Dropdown
ms.assetid: f6a4bb90-be0a-930f-56e7-bc6833af73c3
ms.date: 06/08/2017
---


# ComboBox.Dropdown Method (Access)

You can use the  **Dropdown** method to force the list in the specified combo box to drop down.


## Syntax

 _expression_. **Dropdown**

 _expression_ A variable that represents a **ComboBox** object.


### Return Value

Nothing


## Remarks

For example, you can use this method to cause a combo box listing vendor codes to drop down when the vendor code control receives the focus during data entry.

If the specified combo box control doesn't have the focus, an error occurs. The use of this method is identical to pressing the F4 key when the control has the focus.


## Example

The following example shows how you can use the  **Dropdown** method within the **GotFocus** event procedure to force a combo box named SupplierID to drop down when it receives the focus.


```vb
Private Sub SupplierID_GotFocus() 
 Me!SupplierID.Dropdown 
End Sub
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

