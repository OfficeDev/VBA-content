---
title: Set Properties of Data Access Objects in Visual Basic
keywords: vbaac10.chm5188066
f1_keywords:
- vbaac10.chm5188066
ms.prod: access
ms.assetid: 8942307f-950d-f39d-cab2-ba4fa387b438
ms.date: 06/08/2017
---


# Set Properties of Data Access Objects in Visual Basic



**Applies to:** Access 2013 | Access 2016

Data Access Objects (DAO) enable you to manipulate the structure of your database and the data it contains from Visual Basic. Many DAO objects correspond to objects that you see in your database — for example, a  **TableDef** object corresponds to a Microsoft Access table. A **Field** object corresponds to a field in a table.

Most of the properties you can set for DAO objects are DAO properties. These properties are defined by the Microsoft Access database engine and are set the same way in any application that includes the Access database engine. Some properties that you can set for DAO objects are defined by Microsoft Access, and aren't automatically recognized by the Access database engine. How you set properties for DAO objects depends on whether a property is defined by the Access database engine or by Microsoft Access.

## Setting DAO Properties for DAO Objects

To set a property that's defined by the Access database engine, refer to the object in the DAO hierarchy. The easiest and fastest way to do this is to create object variables that represent the different objects you need to work with, and refer to the object variables in subsequent steps in your code. For example, the following code creates a new  **TableDef** object and sets its **Name** property:


```vb
Dim dbs As DAO.Database 
Dim tdf As DAO.TableDef 
Set dbs = CurrentDb 
Set tdf = dbs.CreateTableDef 
tdf.Name = "Contacts"
```


## Setting Microsoft Access Properties for DAO Objects

When you set a property that's defined by Microsoft Access, but applies to a DAO object, the Access database engine doesn't automatically recognize the property as a valid property. The first time you set the property, you must create the property and append it to the  **Properties** collection of the object to which it applies. Once the property is in the **Properties** collection, it can be set in the same manner as any DAO property.

If the property is set for the first time in the user interface, it's automatically added to the  **Properties** collection, and you can set it normally.

When writing procedures to set properties defined by Microsoft Access, you should include error-handling code to verify that the property you are setting already exists in the  **Properties** collection. See the Help topic about the **CreateProperty** method or the individual property topic for more information.

Keep in mind that when you create the property, you must correctly specify its  **Type** property before you append it to the **Properties** collection. You can determine the **Type** property based on the information in the Settings section of the Help topic for the individual property. The following table provides some guidelines for determining the setting of the **Type** property.



| <strong>If the property setting is</strong>    | <strong>Then the Type property setting should be</strong> |
|:-----------------------------------------------|:----------------------------------------------------------|
| A string                                       | <strong>dbText</strong>                                   |
| <strong>True</strong> / <strong>False</strong> | <strong>dbBoolean</strong>                                |
| An integer                                     | <strong>dbInteger</strong>                                |

The following table lists some Microsoft Access-defined properties that apply to DAO objects.



| <strong>DAO object</strong>             | <strong>Microsoft Access-defined properties</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
|:----------------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Database</strong>               | <strong>AppTitle</strong>, <strong>AppIcon</strong>, <strong>StartupShowDBWindow</strong>, <strong>StartupShowStatusBar</strong>, <strong>AllowShortcutMenus</strong>, <strong>AllowFullMenus</strong>, <strong>AllowBuiltInToolbars</strong>, <strong>AllowToolbarChanges</strong>, <strong>AllowBreakIntoCode</strong>, <strong>AllowSpecialKeys</strong>, <strong>Replicable</strong>, <strong>ReplicationConflictFunction</strong>                                                                                                                                                                                                                            |
| SummaryInfo  <strong>Container</strong> | <strong>Title</strong>, <strong>Subject</strong>, <strong>Author</strong>, <strong>Manager</strong>, <strong>Company</strong>, <strong>Category</strong>, <strong>Keywords</strong>, <strong>Comments</strong>, <strong>Hyperlink Base</strong> (See the <strong>Summary</strong> tab of the ** <em>DatabaseName</em> Properties** dialog box, available by clicking <strong>Database Properties</strong> on the <strong>File</strong> menu.)                                                                                                                                                                                                                     |
| UserDefined  <strong>Container</strong> | (See the  <strong>Summary</strong> tab of the ** <em>DatabaseName</em> Properties** dialog box, available by clicking <strong>Database Properties</strong> on the <strong>File</strong> menu.)                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| <strong>TableDef</strong>               | <strong>DatasheetBackColor</strong>, <strong>DatasheetCellsEffect</strong>, <strong>DatasheetFontHeight</strong>, <strong>DatasheetFontItalic</strong>, <strong>DatasheetFontName</strong>, <strong>DatasheetFontUnderline</strong>, <strong>DatasheetFontWeight</strong>, <strong>DatasheetForeColor</strong>, <strong>DatasheetGridlinesBehavior</strong>, <strong>DatasheetGridlinesColor</strong>, <strong>Description</strong>, <strong>FrozenColumns</strong>, <strong>RowHeight</strong>, <strong>ShowGrid</strong>                                                                                                                                        |
| <strong>QueryDef</strong>               | <strong>DatasheetBackColor</strong>, <strong>DatasheetCellsEffect</strong>, <strong>DatasheetFontHeight</strong>, <strong>DatasheetFontItalic</strong>, <strong>DatasheetFontName</strong>, <strong>DatasheetFontUnderline</strong>, <strong>DatasheetFontWeight</strong>, <strong>DatasheetForeColor</strong>, <strong>DatasheetGridlinesBehavior</strong>, <strong>DatasheetGridlinesColor</strong>, <strong>Description</strong>, <strong>FailOnError</strong>, <strong>FrozenColumns</strong>, <strong>LogMessages</strong>, <strong>MaxRecords</strong>, <strong>RecordLocks</strong>, <strong>RowHeight</strong>, <strong>ShowGrid, UseTransaction</strong> |
| <strong>Field</strong>                  | <strong>Caption</strong>, <strong>ColumnHidden</strong>, <strong>ColumnOrder</strong>, <strong>ColumnWidth</strong>, <strong>DecimalPlaces</strong>, <strong>Description</strong>, <strong>Format</strong>, <strong>InputMask</strong>                                                                                                                                                                                                                                                                                                                                                                                                                            |

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

