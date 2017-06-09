---
title: OptionButton.ColumnOrder Property (Access)
keywords: vbaac10.chm10596
f1_keywords:
- vbaac10.chm10596
ms.prod: access
api_name:
- Access.OptionButton.ColumnOrder
ms.assetid: 5d4d8302-45b4-92e8-4d8f-dc00557ded42
ms.date: 06/08/2017
---


# OptionButton.ColumnOrder Property (Access)

You can use the  **ColumnOrder** property to specify the order of the columns in Datasheet view. Read/write **Integer**.


## Syntax

 _expression_. **ColumnOrder**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

To set or change this property for a table or query by using Visual Basic, you must use a column's  **Properties** collection. For details on using the **Properties** collection, see **Properties**.


 **Note**  The  **ColumnOrder** property isn't available in Design view.

The  **ColumnOrder** property applies to all fields in Datasheet view and to form controls when the form is in Datasheet view.

In Datasheet view, a field's  **ColumnOrder** property setting is determined by the field's position. For example, the field in the leftmost column in Datasheet view has a **ColumnOrder** property setting of 1, the next field has a setting of 2, and so on. Changing a field's **ColumnOrder** property resets the property for that field and every field to the left of its original position in Datasheet view.

In other views, the property setting is 0 unless you explicitly change the order of one or more fields in Datasheet view (either by dragging the fields to new positions or by changing their  **ColumnOrder** property settings). Fields to the right of the moved field's new position will have a property setting of 0 in views other than Datasheet view.

The order of the fields in Datasheet view doesn't affect the order of the fields in table Design view or Form view.


## Example

The following example displays the ProductName and QuantityPerUnit fields in the first two columns in Datasheet view of the Products form.


```vb
Forms!Products!ProductName.ColumnOrder = 1 
Forms!Products!QuantityPerUnit.ColumnOrder = 2
```

The next example displays the ProductName and QuantityPerUnit fields in the first two columns of the Products table in Datasheet view. To set the  **ColumnOrder** property, the example uses the SetFieldProperty procedure. If this procedure is run while the table is open, changes will not be displayed until it is closed and reopened.




```vb
Public Sub SetColumnOrder() 
 
 Dim dbs As DAO.Database 
 Dim tdf As DAO.TableDef 
 
 Set dbs = CurrentDb 
 Set tdf = dbs!Products 
 
 ' Call the procedure to set the ColumnOrder property. 
 SetFieldProperty tdf!ProductName, "ColumnOrder", dbLong, 2 
 SetFieldProperty tdf!QuantityPerUnit, "ColumnOrder", dbLong, 3 
 
 Set tdf = Nothing 
 Set dbs = Nothing 
 
End Sub 
 
Private Sub SetFieldProperty(ByRef fld As DAO.Field, _ 
 ByVal strPropertyName As String, _ 
 ByVal intPropertyType As Integer, _ 
 ByVal varPropertyValue As Variant) 
 ' Set field property without producing nonrecoverable run-time error. 
 
 Const conErrPropertyNotFound = 3270 
 Dim prp As Property 
 
 ' Turn off error handling. 
 On Error Resume Next 
 
 fld.Properties(strPropertyName) = varPropertyValue 
 
 ' Check for errors in setting the property. 
 If Err <> 0 Then 
 If Err <> conErrPropertyNotFound Then 
 On Error GoTo 0 
 MsgBox "Couldn't set property '" &; strPropertyName &; _ 
 "' on field '" &; fld.Name &; "'", vbCritical 
 Else 
 On Error GoTo 0 
 Set prp = fld.CreateProperty(strPropertyName, intPropertyType, _ 
 varPropertyValue) 
 fld.Properties.Append prp 
 End If 
 End If 
 
 Set prp = Nothing 
 
End Sub
```


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

