---
title: Control.Dropdown Method (Access)
keywords: vbaac10.chm10135
f1_keywords:
- vbaac10.chm10135
ms.prod: access
api_name:
- Access.Control.Dropdown
ms.assetid: 45957d42-3e81-f7eb-9579-e5e75c833f59
ms.date: 06/08/2017
---


# Control.Dropdown Method (Access)

You can use the  **Dropdown** method to force the list in the specified combo box to drop down.


## Syntax

 _expression_. **Dropdown**

 _expression_ A variable that represents a **Control** object.


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


[Control Object](control-object-access.md)

