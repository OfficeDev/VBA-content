---
title: ListBox.ItemsSelected Property (Access)
keywords: vbaac10.chm11215
f1_keywords:
- vbaac10.chm11215
ms.prod: access
api_name:
- Access.ListBox.ItemsSelected
ms.assetid: c2403562-00c4-12ec-4d31-9b83d081cb4d
ms.date: 06/08/2017
---


# ListBox.ItemsSelected Property (Access)

You can use the  **ItemsSelected** property to return a read-only reference to the hidden **ItemsSelected** collection. This hidden collection can be used to access data in the selected rows of a multiselect list box control.


## Syntax

 _expression_. **ItemsSelected**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

The  **ItemsSelected** collection is unlike other collections in that it is a collection of **Variants** rather than of objects. Each **Variant** is an integer index referring to a selected row in a list box or combo box.

Use the  **ItemsSelected** collection in conjunction with the **Column** property or the **ItemData** property to retrieve data from selected rows in a list box or combo box. You can list the **ItemsSelected** collection by using the **For Each...Next** statement.

For example, if you have an Employees list box on a form, you can list the  **ItemsSelected** collection and use the control's **ItemData** property to return the value of the bound column for each selected row in the list box.

To enable multiple selection of rows in a list box, set the control's  **MultiSelect** property to Simple or Extended.

The  **ItemsSelected** collection has no methods and two properties, the **Count** and **Item** properties.


## Example

The following example prints the value of the bound column for each selected row in a Names list box on a Contacts form. To try this example, create the list box and set its  **BoundColumn** property as desired and its **MultiSelect** property to Simple or Extended. Switch to Form view, select several rows in the list box, and run the following code:


```vb
Sub BoundData() 
 Dim frm As Form, ctl As Control 
 Dim varItm As Variant 
 
 Set frm = Forms!Contacts 
 Set ctl = frm!Names 
 For Each varItm In ctl.ItemsSelected 
 Debug.Print ctl.ItemData(varItm) 
 Next varItm 
End Sub
```

The next example uses the same list box control, but prints the values of each column for each selected row in the list box, instead of only the values in the bound column.




```vb
Sub AllSelectedData() 
 Dim frm As Form, ctl As Control 
 Dim varItm As Variant, intI As Integer 
 
 Set frm = Forms!Contacts 
 Set ctl = frm!Names 
 For Each varItm In ctl.ItemsSelected 
 For intI = 0 To ctl.ColumnCount - 1 
 Debug.Print ctl.Column(intI, varItm) 
 Next intI 
 Debug.Print 
 Next varItm 
End Sub
```


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

