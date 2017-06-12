---
title: Form.DatasheetFontUnderline Property (Access)
keywords: vbaac10.chm13400
f1_keywords:
- vbaac10.chm13400
ms.prod: access
api_name:
- Access.Form.DatasheetFontUnderline
ms.assetid: a232a1a8-b537-4935-bd64-138548241c7c
ms.date: 06/08/2017
---


# Form.DatasheetFontUnderline Property (Access)

You can use the  **DatasheetFontUnderline** property to specify an underlined appearance for field names and data in Datasheet view. Read/write **Boolean**.


## Syntax

 _expression_. **DatasheetFontUnderline**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **DatasheetFontUnderline** property applies to all fields in Datasheet view and to form controls when the form is in Datasheet view.

This property is only available in [Visual Basic](set-properties-by-using-visual-basic.md)within a Microsoft Access database.

The following table contains the properties that don't exist in the DAO  **Properties** collection of until you set them by using the **Formatting (Datasheet)** toolbar or you can add them in an Access database by using the **CreateProperty** method and append it to the DAO **Properties** collection.


|||
|:-----|:-----|
|**[DatasheetFontItalic](form-datasheetfontitalic-property-access.md)** *|**[DatasheetForeColor](form-datasheetforecolor-property-access.md)** *|
|**[DatasheetFontHeight](form-datasheetfontheight-property-access.md)** *|**[DatasheetBackColor](form-datasheetbackcolor-property-access.md)**|
|**[DatasheetFontName](form-datasheetfontname-property-access.md)** *|**[DatasheetGridlinesColor](form-datasheetgridlinescolor-property-access.md)**|
|**DatasheetFontUnderline** *|**[DatasheetGridlinesBehavior](form-datasheetgridlinesbehavior-property-access.md)**|
|**[DatasheetFontWeight](form-datasheetfontweight-property-access.md)** *|**[DatasheetCellsEffect](form-datasheetcellseffect-property-access.md)**|

 **Note**  When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the  **Properties** collection of the database.


## Example

The following example displays the data and field names in Datasheet view of the Products form as italic and underlined.


```vb
Forms![Products].DatasheetFontItalic = True 
Forms![Products].DatasheetFontUnderline = True
```

The next example displays the data and field names in Datasheet view of the Products table as italic and underlined.

To set the  **DatasheetFontItalic** and **DatasheetFontUnderline** properties, the example uses the SetTableProperty procedure, which is in the database's standard module.




```vb
Dim dbs As Object, objProducts As Object 
Const DB_Boolean As Long = 1 
Set dbs = CurrentDb 
Set objProducts = dbs![Products] 
SetTableProperty objProducts, "DatasheetFontItalic", DB_Boolean, True 
SetTableProperty objProducts, "DatasheetFontUnderline", DB_Boolean, True 
 
Sub SetTableProperty(objTableObj As Object, strPropertyName As String, _ 
 intPropertyType As Integer, varPropertyValue As Variant) 
 ' Set Microsoft Access-defined table property without causing 
 ' nonrecoverable run-time error. 
 Const conErrPropertyNotFound = 3270 
 Dim prpProperty As Variant 
 On Error Resume Next ' Don't trap errors. 
 objTableObj.Properties(strPropertyName) = varPropertyValue 
 If Err <> 0 Then ' Error occurred when value set. 
 If Err <> conErrPropertyNotFound Then 
 On Error GoTo 0 
 MsgBox "Couldn't set property '" &; strPropertyName _ 
 &; "' on table '" &; objTableObj.Name &; "'", 48, "SetTableProperty" 
 Else 
 On Error GoTo 0 
 Set prpProperty = objTableObj.CreateProperty(strPropertyName, _ 
 intPropertyType, varPropertyValue) 
 objTableObj.Properties.Append prpProperty 
 End If 
 End If 
 objTableObj.Properties.Refresh 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

