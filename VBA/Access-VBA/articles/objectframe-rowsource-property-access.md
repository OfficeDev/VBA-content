---
title: ObjectFrame.RowSource Property (Access)
keywords: vbaac10.chm11564
f1_keywords:
- vbaac10.chm11564
ms.prod: access
api_name:
- Access.ObjectFrame.RowSource
ms.assetid: de2aa92d-34e8-20e7-ece7-5e1dcb8cd877
ms.date: 06/08/2017
---


# ObjectFrame.RowSource Property (Access)

You can use the  **RowSource** property (along with the **RowSourceType** property) to tell Microsoft Access how to provide data tothe specified object. Read/write **String**.


## Syntax

 _expression_. **RowSource**

 _expression_ A variable that represents an **ObjectFrame** object.


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


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

