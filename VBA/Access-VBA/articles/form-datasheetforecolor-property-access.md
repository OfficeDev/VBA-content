---
title: Form.DatasheetForeColor Property (Access)
keywords: vbaac10.chm13405
f1_keywords:
- vbaac10.chm13405
ms.prod: access
api_name:
- Access.Form.DatasheetForeColor
ms.assetid: 9756ff09-67bf-edb9-d4b5-d414ec7c1e2a
ms.date: 06/08/2017
---


# Form.DatasheetForeColor Property (Access)

You can use the  **DatasheetForeColor** property in Visual Basic to specify or determine the color of all text in a table, query, or form in Datasheet view within an Access database. Read/write **Long**.


## Syntax

 _expression_. **DatasheetForeColor**

 _expression_ A variable that represents a **Form** object.


## Remarks

Setting the  **DatasheetForeColor** property for a table or query won't affect this property setting for a form that uses the table or query as its source of data.

The following table contains the properties that don't exist in the DAO  **Properties** collection of until you set them by using the **Formatting (Datasheet)** toolbar or you can add them in an Access database by using the **CreateProperty** method and append it to the DAO **Properties** collection.


|||
|:-----|:-----|
|**[DatasheetBackColor](form-datasheetbackcolor-property-access.md)**|**[DatasheetFontUnderline](form-datasheetfontunderline-property-access.md)** *|
|**[DatasheetCellsEffect](form-datasheetcellseffect-property-access.md)**|**[DatasheetFontWeight](form-datasheetfontweight-property-access.md)** *|
|**[DatasheetFontHeight](form-datasheetfontheight-property-access.md)** *|**DatasheetForeColor** *|
|**[DatasheetFontItalic](form-datasheetfontitalic-property-access.md)** *|**[DatasheetGridlinesBehavior](form-datasheetgridlinesbehavior-property-access.md)**|
|**[DatasheetFontName](form-datasheetfontname-property-access.md)** *|**[DatasheetGridlinesColor](form-datasheetgridlinesbehavior-property-access.md)**|

 **Note**   When you add or set any property listed with an asterisk, Microsoft Access automatically adds it to the **Properties** collection.


## Example

The following example uses the SetTableProperty procedure to set a table's font color to dark blue and its background color to light gray. If a "Property not found" error occurs when the property is set, the  **CreateProperty** method is used to add the property to the object's **Properties** collection.


```vb
Dim dbs As Object, objProducts As Object 
Const lngForeColor As Long = 8388608 ' Dark blue. 
Const lngBackColor As Long = 12632256 ' Light gray. 
Const DB_Long As Long = 4 
Set dbs = CurrentDb 
Set objProducts = dbs!Products 
SetTableProperty objProducts, "DatasheetBackColor", DB_Long, lngBackColor 
SetTableProperty objProducts, "DatasheetForeColor", DB_Long, lngForeColor 
 
Sub SetTableProperty(objTableObj As Object, strPropertyName As String, _ 
 intPropertyType As Integer, varPropertyValue As Variant) 
 Const conErrPropertyNotFound = 3270 
 Dim prpProperty As Variant 
 On Error Resume Next ' Don't trap errors. 
 objTableObj.Properties(strPropertyName) = varPropertyValue 
 If Err <> 0 Then ' Error occurred when value set. 
 If Err <> conErrPropertyNotFound Then 
 ' Error is unknown. 
 MsgBox "Couldn't set property '" &; strPropertyName _ 
 &; "' on table '" &; tdfTableObj.Name &; "'", vbExclamation, Err.Description 
 Err.Clear 
 Else 
 ' Error is "Property not found", so add it to collection. 
 Set prpProperty = objTableObj.CreateProperty(strPropertyName, _ 
 intPropertyType, varPropertyValue) 
 objTableObj.Properties.Append prpProperty 
 Err.Clear 
 End If 
 End If 
 objTableObj.Properties.Refresh 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

