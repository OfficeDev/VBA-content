---
title: ListBox.ListCount Property (Access)
keywords: vbaac10.chm11274
f1_keywords:
- vbaac10.chm11274
ms.prod: access
api_name:
- Access.ListBox.ListCount
ms.assetid: 09383f86-888e-1708-9e05-504c49eeb5a6
ms.date: 06/08/2017
---


# ListBox.ListCount Property (Access)

You can use the  **ListCount** property to determine the number of rows in a list box. Read/write **Long**.


## Syntax

 _expression_. **ListCount**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

Microsoft Access sets the  **ListCount** property to the number of rows in the list box or the list box portion of the combo box. The value of the **ListCount** property is read-only and can't be set by the user.

This property is available only by using a macro or Visual Basic . You can read this property only in Form view and Datasheet view.

The  **ListCount** property setting contains the total number of rows in the combo box list or list box, as determined by the control's **RowSource** and **RowSourceType** properties. If the control is based on a table or query (the **RowSourceType** property is set to Table/Query and the **RowSource** property is set to a particular table or query), the **ListCount** property setting contains the number of records in the table or query result set. If the **RowSourceType** property is set to Value List, the **ListCount** property setting contains the number of rows the value list specified in the **RowSource** property results in (this depends on the value list and the number of columns in the list box or combo box list, as set by the **ColumnCount** property).

If you set the  **ColumnHeads** property to Yes, the row of column headings is included in the number of rows returned by the **ListCount** property. For combo boxes and list boxes based on a table or query, adding column headings adds an additional row. For combo boxes and list boxes based on a value list, adding column headings leaves the number of rows unchanged (the first row of values becomes the column headings).

You can use the  **ListCount** property with the **ListRows** property to specify how many rows you want to display in the list box portion of a combo box.


## Example

The following example uses the  **ListCount** property to find the number of rows in the list box portion of the CustomerList combo box on a Customers form. It then sets the **ListRows** property to display a specified number of rows in the list.


```vb
Public Sub SizeCustomerList() 
 
 Dim ListControl As Control 
 
 Set ListControl = Forms!Customers!CustomerList 
 With ListControl 
 If .ListCount < 8 Then 
 .ListRows = .ListCount 
 Else 
 .ListRows = 8 
 End If 
 End With 
 
End Sub
```


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

