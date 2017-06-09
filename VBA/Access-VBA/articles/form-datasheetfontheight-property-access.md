---
title: Form.DatasheetFontHeight Property (Access)
keywords: vbaac10.chm13397
f1_keywords:
- vbaac10.chm13397
ms.prod: access
api_name:
- Access.Form.DatasheetFontHeight
ms.assetid: 5cfcf818-eda0-f7ec-f224-ee52ae7d39c9
ms.date: 06/08/2017
---


# Form.DatasheetFontHeight Property (Access)

You can use the  **DatasheetFontHeight** property to specify the font point size used to display and print field names and data in Datasheet view. Read/write **Integer**.


## Syntax

 _expression_. **DatasheetFontHeight**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is only available within a Microsoft Access database.

For the  **DatasheetFontHeight** property, the font size you specify must be valid for the font specified by the **DatasheetFontName** property. For example, MS Sans Serif is available only in sizes 8, 10, 12, 14, 18, and 24 points.

The following table contains the properties that don't exist in the DAO  **Properties** collection of until you set them by using the **Formatting (Datasheet)** toolbar or you can add them in an Access database (.mdb) by using the **CreateProperty** method and append it to the **DAO Properties** collection.


|||
|:-----|:-----|
|**[DatasheetFontItalic](form-datasheetfontitalic-property-access.md)** *|**[DatasheetForeColor](form-datasheetforecolor-property-access.md)** *|
|**[DatasheetFontHeight](form-datasheetfontheight-property-access.md)** *|**[DatasheetBackColor](form-datasheetbackcolor-property-access.md)**|
|**[DatasheetFontName](form-datasheetfontname-property-access.md)** *|**[DatasheetGridlinesColor](form-datasheetgridlinescolor-property-access.md)**|
|**[DatasheetFontUnderline](form-datasheetfontunderline-property-access.md)** *|**[DatasheetGridlinesBehavior](form-datasheetgridlinesbehavior-property-access.md)**|
|**[DatasheetFontWeight](form-datasheetfontweight-property-access.md)** *|**[DatasheetCellsEffect](form-datasheetcellseffect-property-access.md)**|

 **Note**  When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the  **Properties** collection of the database.


## Example

The following example sets the font to MS Serif, the font size to 10 points, and the font weight to medium (500) in Datasheet view of the Products table.


```vb
Sub SetDatasheetFont 
 
 Dim dbs As Object, objProducts As Object 
 Set dbs = CurrentDb 
 Const DB_Text As Long = 10 
 Const DB_Integer As Long = 3 
 Set objProducts = dbs!Products 
 
 SetTableProperty objProducts, "DatasheetFontName", DB_Text, "MS Serif" 
 SetTableProperty objProducts, "DatasheetFontHeight", DB_Integer, 10 
 SetTableProperty objProducts, "DatasheetFontWeight", DB_Integer, 500 
 
End Sub 
 
 
 
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

The next example makes the same changes as the preceding example in Datasheet view of the open Products form.




```vb
Forms!Products.DatasheetFontName = "MS Serif" 
Forms!Products.DatasheetFontHeight = 10 
Forms!Products.DatasheetFontWeight = 500
```


## See also


#### Concepts


[Form Object](form-object-access.md)

