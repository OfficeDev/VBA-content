---
title: ComboBox.Column Property (Access)
keywords: vbaac10.chm11360
f1_keywords:
- vbaac10.chm11360
ms.prod: access
api_name:
- Access.ComboBox.Column
ms.assetid: 3b410a44-9055-e2c7-b921-4b364f68041b
ms.date: 06/08/2017
---


# ComboBox.Column Property (Access)

You can use the  **Column** property to refer to a specific column, or column and row combination, in a multiple-column combo box or list box. Read-only **Variant**.


## Syntax

 _expression_. **Column**( ** _Index_**, ** _Row_** )

 _expression_ A variable that represents a **ComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|A long integer that can range from 0 to the setting of the  **ColumnCount** property minus one.|
| _Row_|Optional|**Variant**|An integer that can range from 0 to the setting of the  **ListCount** property minus 1.|

## Remarks

Use 0 to refer to the first column, 1 to refer to the second column, and so on. Use 0 to refer to the first row, 1 to refer to the second row, and so on. For example, in a list box containing a column of customer IDs and a column of customer names, you could refer to the customer name in the second column and fifth row as:


```vb
Forms!Contacts!Customers.Column(1, 4)
```

You can use the  **Column** property to assign the contents of a combo box or list box to another control, such as a text box. For example, to set the **ControlSource** property of a text box to the value in the second column of a list box, you could use the following expression:




```text
=Forms!Customers!CompanyName.Column(1)
```

If the user has made no selection when you refer to a column in a combo box or list box, the  **Column** property setting will be **Null**. You can use the **IsNull** function to determine if a selection has been made, as in the following example:




```vb
If IsNull(Forms!Customers!Country) 
  Then MsgBox "No selection." 
End If
```


 **Note**  To determine how many columns a combo box or list box has, you can inspect the  **ColumnCount** property setting.


## Example

The following example uses the  **Column** property and the **ColumnCount** property to print the values of a list box selection.


```vb
Public Sub Read_ListBox() 
 
 Dim intNumColumns As Integer 
 Dim intI As Integer 
 Dim frmCust As Form 
 
 Set frmCust = Forms!frmCustomers 
 If frmCust!lstCustomerNames.ItemsSelected.Count > 0 Then 
 
 ' Any selection? 
 intNumColumns = frmCust!lstCustomerNames.ColumnCount 
 Debug.Print "The list box contains "; intNumColumns; _ 
 IIf(intNumColumns = 1, " column", " columns"); _ 
 " of data." 
 
 Debug.Print "The current selection contains:" 
 For intI = 0 To intNumColumns - 1 
 ' Print column data. 
 Debug.Print frmCust!lstCustomerNames.Column(intI) 
 Next intI 
 Else 
 Debug.Print "You haven't selected an entry in the " _ 
 &; "list box." 
 End If 
 
 Set frmCust = Nothing 
 
End Sub
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

