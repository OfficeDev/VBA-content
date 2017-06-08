---
title: CheckBox.ValidationRule Property (Access)
keywords: vbaac10.chm10698
f1_keywords:
- vbaac10.chm10698
ms.prod: access
api_name:
- Access.CheckBox.ValidationRule
ms.assetid: 4ebb1371-acd0-2227-49e9-ec646a0daaad
ms.date: 06/08/2017
---


# CheckBox.ValidationRule Property (Access)

You can use the  **ValidationRule** property to specify requirements for data entered into a record, field, or control. When data is entered that violates the **ValidationRule** setting, you can use the **ValidationText** property to specify the message to be displayed to the user. Read/write **String**.


## Syntax

 _expression_. **ValidationRule**

 _expression_ A variable that represents a **CheckBox** object.


## Remarks

Enter an expression for the  **ValidationRule** property setting and text for the **ValidationText** property setting. The maximum length for the **ValidationRule** property setting is 2048 characters. The maximum length for the **ValidationText** property setting is 255 characters.

For controls, you can set the  **ValidationRule** property to any valid expression. For field and record validation rules, the expression can't contain user-defined functions, domain aggregate or aggregate functions, the **Eval** function, or **CurrentUser** method, or references to forms, queries, or tables. In addition, field validation rules can't contain references to other fields. For records, expressions can include references to fields in that table.

For table fields and records, you can also set these properties in Visual Basic by using the DAO  **ValidationRule** property.


 **Note**  The  **ValidationRule** and **ValidationText** properties don't apply to check box, option button, or toggle buttoncontrols when they are in an option group. They apply only to the option group itself.

Microsoft Access automatically validates values based on a field's data type; for example, Microsoft Access doesn't allow text in a numeric field. You can set rules that are more specific by using the  **ValidationRule** property.

If you set the  **ValidationRule** property but not the **ValidationText** property, Microsoft Access displays a standard error message when the validation rule is violated. If you set the **ValidationText** property, the text you enter is displayed as the error message.

For example, when a record is added for a new employee, you can enter a  **ValidationRule** property requiring that the value in the employee's StartDate field fall between the company's founding date and the current date. If the date entered isn't in this range, you can display the **ValidationText** property message: "Start date is incorrect."

If you create a control by dragging a field from the field list, the field's validation rule remains in effect, although it isn't displayed in the control's  **ValidationRule** property box in the property sheet. This is because a field's validation rule is inherited by a control bound to that field.

Control, field, and record validation rules are applied as follows:


- Validation rules you set for fields and controls are applied when you edit the data and the focus leaves the field or control.
    
- Validation rules for records are applied when you move to another record.
    
- If you create validation rules for both a field and a control bound to the field, both validation rules are applied when you edit data and the focus leaves the control.
    
The following table contains expression examples for the  **ValidationRule** and **ValidationText** properties.



|**ValidationRule property**|**ValidationText property**|
|:-----|:-----|
|<> 0|Entry must be a nonzero value.|
|> 1000 Or Is Null|Entry must be blank or greater than 1000.|
|Like "A????"|Entry must be 5 characters and begin with the letter "A".|
|>= #1/1/96# And <#1/1/97#|Entry must be a date in 1996.|
|DLookup("CustomerID", "Customers", "CustomerID = Forms!Customers!CustomerID") Is Null|Entry must be a unique CustomerID (domain aggregate functions are allowed only for form-level validation).|
If you create a validation rule for a field, Microsoft Access doesn't normally allow a  **Null** value to be stored in the field. If you want to allow a **Null** value, add "Is Null" to the validation rule, as in "<> 8 Or Is Null" and make sure the **Required** property is set to No.

You can't set field or record validation rules for tables created outside Microsoft Access (for example, dBASE, Paradox, or SQL Server). For these kinds of tables, you can create validation rules for controls only.


## Example

The following example creates a validation rule for a field that allows only values over 65 to be entered. If a number less than 65 is entered, a message is displayed. The properties are set by using the SetFieldValidation function.


```vb
Dim strTblName As String, strFldName As String 
Dim strValidRule As String 
Dim strValidText As String, intX As Integer 
 
strTblName = "Customers" 
strFldName = "Age" 
strValidRule = ">= 65" 
strValidText = "Enter a number greater than or equal to 65." 
intX = SetFieldValidation(strTblName, strFldName, _ 
 strValidRule, strValidText) 
 
Function SetFieldValidation(strTblName As String, _ 
 strFldName As String, strValidRule As String, _ 
 strValidText As String) As Integer 
 
 Dim dbs As Database, tdf As TableDef, fld As Field 
 
 Set dbs = CurrentDb 
 Set tdf = dbs.TableDefs(strTblName) 
 Set fld = tdf.Fields(strFldName) 
 fld.ValidationRule = strValidRule 
 fld.ValidationText = strValidText 
End Function
```

The next example uses the SetTableValidation function to set record-level validation to ensure that the value in the EndDate field comes after the value in the StartDate field.




```vb
Dim strTblName As String, strValidRule As String 
Dim strValidText As String 
Dim intX As Integer 
 
strTblName = "Employees" 
strValidRule = "EndDate > StartDate" 
strValidText = "Enter an EndDate that is later than the StartDate." 
intX = SetTableValidation(strTblName, strValidRule, strValidText) 
 
Function SetTableValidation(strTblName As String, _ 
 strValidRule As String, strValidText As String) _ 
 As Integer 
 
 Dim dbs As Database, tdf As TableDef 
 
 Set dbs = CurrentDb 
 Set tdf = dbs.TableDefs(strTblName) 
 tdf.ValidationRule = strValidRule 
 tdf.ValidationText = strValidText 
End Function
```


## See also


#### Concepts


[CheckBox Object](checkbox-object-access.md)

