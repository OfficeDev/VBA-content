---
title: Set Properties of ActiveX Data Objects in Visual Basic
ms.prod: access
ms.assetid: 54955634-d354-54ff-495b-1f696e392dfe
ms.date: 06/08/2017
---


# Set Properties of ActiveX Data Objects in Visual Basic

ActiveX Data Objects (ADO) enable you to manipulate the structure of your database and the data it contains from Visual Basic. Many ADO objects correspond to objects that you see in your database—for example, a  **Table** object corresponds to an Access table. A **Field** object corresponds to a field in a table.

Most of the properties you can set for ADO objects are ADO properties. These properties are defined by the Access database engine and are set the same way in any application that includes the Access database engine. Some properties that you can set for ADO objects are defined by Access, and are not automatically recognized by the Access database engine. How you set properties for ADO objects depends on whether a property is defined by the Access database engine or by Access.

## Setting ADO Properties for ADO Objects

To set a property that is defined by the Access database engine, refer to the object in the ADO hierarchy. The easiest and fastest way to do this is to create object variables that represent the different objects you need to work with, and refer to the object variables in subsequent steps in your code. For example, the following code creates a new  **TableDef** object and sets its **Name** property:


```vb
Dim tbl As New ADOX.Table 
Dim cnn As ADODB.Connection 
Set cnn = CurrentProject.Connection 
tbl.Name = "Contacts"
```


## Setting Access Properties for ADO Objects

When you set a property that is defined by Access, but applies to an ADO object, the Access database engine does not automatically recognize the property as a valid property. The first time you set the property, you must create the property and append it to the  **Properties** collection of the object to which it applies. Once the property is in the **Properties** collection, it can be set in the same manner as any ADO property.

If the property is set for the first time in the user interface, it is automatically added to the  **Properties** collection, and you can set it normally.

When writing procedures to set properties defined by Access, you should include error-handling code to verify that the property you are setting already exists in the  **Properties** collection.

Keep in mind that when you create the property, you must correctly specify its  **Type** property before you append it to the **Properties** collection. You can determine the **Type** property based on the information in the Settings section of the Help topic for the individual property. The following table provides some guidelines for determining the setting of the **Type** property.



|**If the property setting is**|**Then the Type property setting should be**|
|:-----|:-----|
|A string|**adLongVarWChar** or **adVarWChar**|
|**True** / **False**|**adBoolean**|
|An integer|**adInteger**|
The following table lists some Access-defined properties that apply to ADO objects.



|**ADO object**|**Microsoft Access-defined properties**|
|:-----|:-----|
|**Connection**|**AppTitle**, **AppIcon**, **StartupShowDBWindow**, **StartupShowStatusBar**, **AllowShortcutMenus**, **AllowFullMenus**, **AllowBuiltInToolbars**, **AllowToolbarChanges**, **AllowBreakIntoCode**, **AllowSpecialKeys**, **Replicable**, **ReplicationConflictFunction**|
|**Table**|**DatasheetBackColor**, **DatasheetCellsEffect**, **DatasheetFontHeight**, **DatasheetFontItalic**, **DatasheetFontName**, **DatasheetFontUnderline**, **DatasheetFontWeight**, **DatasheetForeColor**, **DatasheetGridlinesBehavior**, **DatasheetGridlinesColor**, **Description**, **FrozenColumns**, **RowHeight**, **ShowGrid**|
|**Field**|**Caption**, **ColumnHidden**, **ColumnOrder**, **ColumnWidth**, **DecimalPlaces**, **Description**, **Format**, **InputMask**|

