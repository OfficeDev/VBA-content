---
title: ComboBox.BoundColumn Property (Access)
keywords: vbaac10.chm11383
f1_keywords:
- vbaac10.chm11383
ms.prod: access
api_name:
- Access.ComboBox.BoundColumn
ms.assetid: ba2b5807-5f5a-52bb-d5d3-db7525bccba4
ms.date: 06/08/2017
---


# ComboBox.BoundColumn Property (Access)

When you make a selection from a combo box, the  **BoundColumn** property tells Microsoft Access which column's values to use as the value of the control. If the control is bound to a field, the value in the column specified by the **BoundColumn** property is stored in the field named in the **ControlSource** property. Read/write **Long**.


## Syntax

 _expression_. **BoundColumn**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **BoundColumn** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|The  **ListIndex** property value, rather than the column value, is stored in the current record. The **ListIndex** property value of the first row is 0, the second row is 1, and so on. Microsoft Access sets the **ListIndex** property when an item is selected from a list box or the list box portion of a combo box. Setting the **BoundColumn** property to 0 and using the **ListIndex** property value of the control might be useful if, for example, you are only interested in storing a sequence of numbers.|
|1 or greater|(Default is 1) The value in the specified column becomes the control's value. If the control is bound to a field, then this setting is stored in that field in the current record. The BoundColumn property can't be set to a value larger than the setting of the ColumnCount property.|
For table fields , you can set this property on the  **Lookup** tab in the Field Properties section of table Design view for fields with the **DisplayControl** property set to Combo Box or List Box.

In Visual Basic, set the  **BoundColumn** property by using a number or a numeric expression equal to a value from 0 to the setting of the **ColumnCount** property.

The leftmost visible column in a combo box (the leftmost column whose setting in the combo box's  **ColumnWidths** property is not 0) contains the data that appears in the text box part of the combo box in Form view or in a report. The **BoundColumn** property determines which column's value in the text box or combo box list will be stored when you make a selection. This allows you to display different data than you store as the value of the control.


 **Note**  If the bound column is not the same as the leftmost visible column in the control (or if you set the  **BoundColumn** property to 0), the **LimitToList** property is set to Yes.

Microsoft Access uses zero-based numbers to refer to columns in the  **Column** property. That is, the first column is referenced by using the expression `Column(0)`; the second column is referenced by using the expression  `Column(1)`; and so on. However, the  **BoundColumn** property uses 1-based numbers to refer to the columns. This means that if the **BoundColumn** property is set to 1, you could access the value stored in that column by using the expression `Column(0)`.

If the  **AutoExpand** property is set to Yes, Microsoft Access automatically fills in a value in the text box portion of the combo box that matches a value in the combo box list as you type.


## Example

The following example show how to create a combo box that is bound to one column while displaying another. Setting the  **ColumnCount** property to 2 specifies that the **cboDept** combo box will display the first two columns of the data source specified by the **RowSource** property. Setting the **BoundColumn** property to 1 specifies that the value stored in the first column will be returned when you inspect the value of the combo box.

The  **ColumnWidths** property specifies the width of the two columns. By setting the width of the first column to **0in.**, the first column is not displayed in the combo box.

 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```vb
Private Sub cboDept_Enter()
    With cboDept
        .RowSource = "SELECT * FROM tblDepartments ORDER BY Department"
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "0in.;1in."
    End With
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[ComboBox Object](combobox-object-access.md)

