---
title: ComboBox.RowSource Property (Access)
keywords: vbaac10.chm11379
f1_keywords:
- vbaac10.chm11379
ms.prod: access
api_name:
- Access.ComboBox.RowSource
ms.assetid: 1225e566-24e0-244d-09ae-e036c87f3141
ms.date: 06/08/2017
---


# ComboBox.RowSource Property (Access)

You can use the  **RowSource** property (along with the **RowSourceType** property) to tell Microsoft Access how to provide data tothe specified object. Read/write **String**.


## Syntax

 _expression_. **RowSource**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

For example, to display rows of data in a list box from a query named CustomerList, set the list box's  **RowSourceType** property to Table/Query and its **RowSource** property to the query named CustomerList.

The  **RowSource** property setting depends on the **RowSourceType** property setting.



|**For this RowSourceType setting**|**Enter this RowSource setting**|
|:-----|:-----|
|Table/Query|A table name, query name, or SQL statement.|
|Value List|A list of items with semicolons (;) as separators.|
|Field List|A table name, query name, or SQL statement.|
If the  **RowSourceType** property is set to a user-defined function, the **RowSource** property can be left blank.

For table fields, you can set these properties on the  **Lookup** tab in the Field Properties section of table Design view for fields with the **DisplayControl** property set to Combo Box or List Box.

Microsoft Access sets these properties automatically when you select Lookup Wizard as the data type for a field in table Design view.

In Visual Basic, set the  **RowSourceType** property by using a string expression with one of these values: `"Table/Query"`,  `"Value List"`, or  `"Field List"`. You also use a string expression to set the value of the  **RowSource** property. To set the **RowSourceType** property to a user-defined function, enter the name of the function.

When you have a limited number of values that don't change, you can set the  **RowSourceType** property to Value List and then enter the values that make up the list in the **RowSource** property.


## Example

The following example sets the  **RowSourceType** property for a combo box to Table/Query, and it sets the **RowSource** property to a query named EmployeeList.


```vb
Forms!Employees!cmboNames.RowSourceType = "Table/Query" 
Forms!Employees!cmboNames.RowSource = "EmployeeList"
```



The following example shows how to set the  **RowSource** property of a combo box when a form is loaded. When the form is displayed, the items stored in the **Departments** field of the **tblDepartment** combo box are displayed in the **cboDept** combo box.

 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```vb
Private Sub Form_Load()
    Me.Caption = "Today is " &; Format$(Date, "dddd mmm-d-yyyy")
    Me.RecordSource = "tblDepartments"
    DoCmd.Maximize  
    txtDept.ControlSource = "Department"
    cmdClose.Caption = "&;Close"
    cboDept.RowSourceType = "Table/Query"
    cboDept.RowSource = "SELECT Department FROM tblDepartments"
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[ComboBox Object](combobox-object-access.md)

