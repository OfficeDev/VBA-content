---
title: About the PowerPivot Model Object in Excel
ms.prod: excel
ms.assetid: baa95a62-53d2-4c5f-bff7-bcc7323d6a20
ms.date: 06/08/2017
---


# About the PowerPivot Model Object in Excel
Learn about the PowerPivot add-in model and its object model in Excel.

## About the PowerPivot Model Object
<a name="XLPowerPivotModel_About"> </a>

The PowerPivot add-in enables you to visually build your own cubes. A data cube is an array of data defined in dimensions or layers. The  **Model** object in Excel implemented by the PowerPivot add-in provides the foundation to load and combine source data from several data sources for data analysis on the desktop, including relational databases, multidimensional sources, cloud services, data feeds, Excel files, text files, and data from the Web. Excel integrates additional data sources and enables the ability to combine data from multiple data sources.

The creation and deletion of the PowerPivot Model (PPM) is triggered by user exposed actions and cannot be created directly by the developer.


## Relationships defined
<a name="XLPowerPivotModel_Relationships"> </a>

Throughout this article, we will refer to the connection between two tables that establishes how the data should be correlated as relationships.

Relationships join together data from previously unrelated data sources. Each relationship has a  _Primary Key_ and a _Foreign Key_. Relationships allow the data to be joined together into a single model. This allows for:


- Filtering data in one table by data in a related table

- Filtering data by related columns
    
- Integrating columns from multiple tables into a PivotTable/PivotChart
    
- Keeping workbooks smaller by not having to repeat data
    

## Single Models Only
<a name="XLPowerPivotModel_Single"> </a>

Excel with the PowerPivot add-in creates a single model in the workbook to which it can add data sources, create, modify, and relate tables. There can only be a single model in a workbook.


## Working with OLAP data sources
<a name="XLPowerPivotModel_OLAP"> </a>

When connecting to an OLAP data source such as Analysis Services and creating OLAP PivotTables, Pivot Charts, Slicers or Cube functions, no model is created. Workbooks created with the PowerPivot add-in can be uploaded to SharePoint, loaded in memory on the server, and accessed by other workbooks as if it were a normal instance of SQL Server Analysis Services.


## Triggering the creation of a PowerPivot Model
<a name="XLPowerPivotModel_Triggering"> </a>

By default, XLSX files in Excel 2010 and Excel do not have a PPM initialized in them until the model is deemed necessary. Certain actions trigger the creation of a PPM if there is no existing model in the workbook. The following sections describe the actions that will trigger the creation of a PPM when it does not exist in the workbook.


### Adding a new non-legacy data source

Any time you import certain types of data, a new model is created in the workbook (if one does not already exist) that contains the connection properties, table representation of the workbook data sources, and the relationships between them. This includes internal data sources like ranges and tables. Table 1 lists the different data sources that can be integrated with the PPM.

 **Table 1. Data sources compatible with the PowerPivot Model**



|**Data Source**|**Description**|**Table Preview**|**Query Supported**|
|:-----|:-----|:-----|:-----|
|Microsoft SQL Server|Already supported in Excel|Yes|Yes|
|Microsoft SQL Azure Data Market|Supported as a new data feed data source|Yes|No|
|Microsoft SQL Server Parallel Data Warehouse|Supported via installed OLE DB driver|Yes|Yes|
|Microsoft Access|Already supported in Excel|Yes|Yes|
|Oracle|Already supported in Excel|Yes|Yes|
|Teradata|Available if OLE DB or ODBC driver is installed|No|No|
|Sybase|Available if OLE DB or ODBC driver is installed|No|No|
|Informix|Available if OLE DB or ODBC driver is installed|No|No|
|IBM Db2|Available if OLE DB or ODBC driver is installed|No|No|
|Microsoft Analysis Services|Already supported in Excel|Yes|Yes|
|Report (SSRS)|Can read and use connections, but no authoring in Excel client|Yes|No|
|Text|From Excel dialog in Ribbon UI|Yes|No|
|Data Feeds (OData)|Supported as a new data source|Yes|Yes|
|XML|Already supported in Excel|No|No|
|SharePoint Lists|Already supported in Excel. Excel uses the  **DataFeed** provider to connect to SharePoint|No|No|
|SharePoint|New feature in Excel|Yes|Yes|
|Excel Tables|User defined table in Excel used for new data feature. A Worksheet data connection is created to the table when the table is created.|N/A|N/A|
|Excel Ranges|User defined range in Excel used for new data feature. A Worksheet data connection in this case is created to the range only if a data feature like a chart or PivotTable uses the range.|N/A|N/A|

### Creating a new Excel non-OLAP PivotTable

New Excel PivotTables, other than the ones created from an OLAP data source, will be based on a PPM therefore if a PPM is not present in the file a new one is created as part of the PivotTable creation action. This includes the following:


- Using the insert PivotTable user interface
    
- Summarizing data with PivotTable user interface
    
- PivotTable based off of a non-OLAP data source created through the Microsoft Visual Basic for Applications (VBA) object model
    

### Creating a new Excel non-OLAP PivotChart

In Excel, PivotTables and Pivot Charts have the ability to be no longer coupled. Therefore on insertion of a PivotChart in a workbook without a model, a PPM will be created.


### Pasting Excel non-OLAP PivotTables from another workbook

When pasting a PivotTable or PivotChart from another workbook that is based off of a PPM into one that does not have a PPM, a new PPM will be created in the destination workbook. A new data source will be added to the newly created model pointing to the underlying data of the originating PivotTable/PivotChart.


## Undoing the creation of a PowerPivot Model
<a name="XLPowerPivotModel_Undoing"> </a>

All actions that lead to the creation of a PPM can be undone. If these actions are selected from the undo menu, the actual model creation will not be undone but nothing will be added to it; therefore it will remain empty. When the workbook is saved, if the model is empty, the model will not be saved with the file. There is no explicit way for you to manually delete a model created in the workbook.


 **Note**  Similar to the behavior in Excel 2010, there is a restriction in what model sizes can be undone. When a model grows to this limit size undo functionality for actions such as refresh will no longer be provided. The current limit for native PivotTables is 300,000 rows, at 28 bytes a cell this limit is roughly 8MB in memory. These values can be set by using  **Advanced Options** in Excel as shown in Figure 1.


**Figure 1. Set the size for large Data Model undo operations**

![Set size of Data Model undo operations](images/XLPowerPivotModel_01.jpg)


## The PowerPivot Model Object Model
<a name="XLPowerPivotModel_PPMOM"> </a>

A workbook will be able to have one and only one  **Model** object. The **Model** object represents the top level object that contains all its connections, relationships, and tables.

You are not able to manually create a model in a workbook; creation of the model is triggered through the actions described in a previous section in this article. If any of these actions are performed through the Object Model (OM), a new model is created. The purpose of this OM is for the programmatic creation of relationships between model tables resulting in joined tables, combining PivotTables, and so forth. For you to be able to this, you must be able to explore the model to find the appropriate tables and within the tables find the appropriate columns that would be used to create the relationship.


### Model Object

The  **Model** object stores references to workbook connections and information about the tables and relationships contained within the PPM. Table 2 lists the properties of the **Model** object.

 **Table 2. Properties of the Model object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **Model** object.|
| **ModelTables**|Read-only| **ModelTable**|Collection of tables inside the PPM.|
| **ModelRelationships**|Read-only| **ModelRelationships**|Collection of relationships between PPM tables.|
| **DataModelConnection**|N/A| **WorkbookConnection**|Returns the model workbook connection object from the workbook connections collection which connects to the model.|
 **Model.AddConnection** Method

Adds a new workbook connection to the model with the same properties as the one supplied as an argument. This method only works on non-model external connections and will return an error if called with an external model connection as its argument. When calling this method, a new model connection is created and it is named the same as the legacy connection with an integer at the end to make the name unique. Table 3 lists the parameters of the  **AddConnection** method.

 **Table 3. Parameters of the Model.AddConnection method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ConnectionToDataSource|Required| **WorkbookConnection**|The Workbook connection|
 **Model.CreateModelWorkbookConnection** Method

Calling this method returns a  **WorkbookConnection** object of type **ModelConnection**. A model connection connected to the specified table is returned. This type of connection can only be used by query tables in Excel. Table 4 lists the parameters of the  **CreateModelWorkbookConnection** method.

 **Table 4. Parameters of the Model.CreateModelWorkbookConnection method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ModelTable|Required| **Variant**|Either a model table name or a model table object.|
 **Model.Initialize** Method

The  **Initialize** method of the **Model** object has no parameters. Initializes the PPM. This is called by default the first time the model is used.

 **Model.Refresh** Method

The  **Refresh** method of the **Model** object has no parameters. Refreshes all data sources associated with the model, fully reprocesses the model and updates all Excel data features associated with the **Model** object.


### ModelChanges Object

Represents changes made to the PPM. The  **ModelChanges** object contains information about which changes were made to the data model when the **Workbook.ModelChange** event occurs after a model operation. When Excel makes changes to the data model, multiple changes can be made in the same operation and the **ModelChanges** object will include information about all the changes made in one model operation. Table 5 lists the properties of the **ModelChanges** object.

 **Table 5. Properties of the ModelChanges object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **ColumnsAdded**|Read-only| **ModelColumnNames**|Returns a  **ModelColumnNames** collection of **ModelColumnName** objects which represent all columns added as part of a model operation.|
| **ColumnsChanged**|Read-only| **ModelColumnChanges**|Returns a  **ModelColumnChanges** collection of **ModelColumnChange** objects which represent table names and column names of all table columns for which the data type was changed as part of a model operation.|
| **ColumnsDeleted**|Read-only| **ModelColumnNames**|Returns a  **ModelColumnNames** collection of **ModelColumnName** objects which represent all columns which were deleted as part of a model operation.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **MeasuresAdded**|Read-only| **ModelMeasureNames**|Returns a  **ModelMeasureNames** collection of **ModelMeasureName** objects which represent all measures which were added as part of a model operation.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelChanges** object.|
| **RelationshipChange**|Read-only| **Boolean**|When  **True**, one or more relationships in the model were changed (added, deleted or modified) as part of a model operation. When  **False**, no relationships were changed during the operation.|
| **TableNamesChanged**|Read-only| **ModelTableNameChanges**|Returns a  **ModelTableNameChanges** collection of **ModelTableNameChange** objects that represents old and new names of all tables which were renamed in the model as part of a model operation.|
| **TablesAdded**|Read-only| **ModelTableNames**|Returns a  **ModelTableNames** collection of table names as strings that represents all tables which were added to the model as part of a model operation.|
| **TablesDeleted**|Read-only| **ModelTableNames**|Returns a  **ModelTableNames** collection of table names as strings that represents all tables which were deleted from the model as part of a model operation.|
| **TablesModified**|Read-only| **ModelTableNames**|Returns a  **ModelTableNames** collection of table names as strings that represents all tables which were refreshed or recalculated as part of a model operation.|
| **UnknownChange**|Read-only| **Boolean**| **True** when a non-specified change was made to the model as part of a model transaction.|

### ModelColumnChanges Collection

A collection of  **ModelColumnChange** objects that represent columns for which the data type was change in the PPM. Table 6 lists the properties of the **ModelColumnChanges** collection.

 **Table 6. Properties of the ModelColumnChanges collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelColumnChange** objects in the collection|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelColumnChanges** object.|
 **ModelColumnChanges.Item** Method

Returns a single object from the  **ModelColumnChanges** collection. Table 7 lists the parameters of the **Item** method.

 **Table 7. Parameters of the ModelColumnChanges.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelColumnChange Object

An object that represents a column in a table in the PPM for which the data type was changed. Table 8 lists the properties of the  **ModelColumnChange** object.

 **Table 8. Properties of the ModelColumnChange object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **ColumnName**|Read-only| **String**| **String** that represents the name of a column for which the data type was changed.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelColumnChange** object.|
| **TableName**|Read-only| **String**| **String** that represents the name of a table in the PPM for which the data type of a column was changed.|

### ModelColumnNames Collection

A collection of  **ModelColumnName** objects that represents columns of tables in the PPM. Table 9 lists the properties of the **ModelColumnNames** collection.

 **Table 9. Properties of the ModelColumnNames collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelColumnName** objects in the collection|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelColumnNames** collection.|
 **ModelColumnNames.Item** Method

Returns a single object from the  **ModelColumnNames** collection. Table 10 lists the parameters of the **Item** method

 **Table 10. Parameters of the ModelColumnNames.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelColumnName Object

An object that represents the name of a column in the PPM. Table 11 lists the properties of the  **ModelColumnName** object.

 **Table 11. Properties of the ModelColumnName object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **ColumnName**|Read-only| **String**| **String** that represents the name of a column of the table identified by the **TableName** property.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelColumnName** object.|
| **TableName**|Read-only| **String**| **String** that represents the name of a table in the PPM.|

### ModelConnection Object

The  **ModelConnection** object will contain information for the new Model Connection Type introduced in Excel to interact with the integrated PPM. Table 12 lists the properties of the **ModelConnection** object.

 **Table 12. Properties of the ModelConnection object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **ADOConnection**|Read-only| **ADOConnection**|Used to create an open connection to a data source. Enables add-ins, such as the PowerViewer, to create a direct connection to the engine and hence the data model.|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **CommandText**|Read/Write| **Variant**|Returns or sets the command string for the specified data source (table).|
| **CommandType**|Read/Write| **xlCmdType**|Returns or sets one of the  **xlCmdType** constants specifying the command type.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelConnection** object.|

### ModelMeasureNames Collection

The  **ModelMeasureNames** collection contains a collection of **ModelMeasureName** objects in the PPM. Table 13 lists the properties of the **ModelMeasureNames** collection.

 **Table 13. Properties of the ModelMeasureNames collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelMeasureName** objects in the collection|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelMeasureNames** collection.|
 **ModelMeasureNames.Item** Method

Returns a single object from the  **ModelMeasureNames** collection. Table 14 list the parameters of the **Item** method.

 **Table 14. Parameters of the ModelMeasureNames.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelMeasureName Object

An object that represents the name of a measure in the PPM. Table 15 lists the properties of the  **ModelMeasureName** object.

 **Table 15. Properties of the ModelMeasureName object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **MeasureName**|Read-only| **String**| **String** that represents the new name a measure which was added to the **ModelTable** object identified by the **TableName** property.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelMeasureName** object.|
| **TableName**|Read-only| **String**| **String** that represents the name of a table in the PPM.|

### ModelRelationships Collection

The  **ModelRelationships** collection contains a collection of **ModelRelationship** objects in the PPM. Table 16 lists the properties of the **ModelRelationships** collection.

 **Table 16. Properties of the ModelRelationships collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelRelationship** objects in the collection|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelRelationships** collection.|
 **ModelRelationships.Add** Method

Adds a relationship to the  **ModelRelationships** collection. Table 17 lists the parameters of the **Add** method.

 **Table 17. Parameters of the ModelRelationships.Add method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ForeignKeyColumn|Required| **ModelTableColumn**|A  **ModelTableColumn** object that represents the foreign key column in the table on the many side of the one-to-many relationship.|
|PrimaryKeyColumn|Required| **ModelTableColumn**|A  **ModelTableColumn** object that represents the primary key column in the table on the one side of the one-to-many relationship.|
 **ModelRelationships.Item** Method

Returns a single object from the  **ModelRelationships** collection. Table 18 lists the parameters of the **Item** method.

 **Table 18. Parameters of the ModelRelationships.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelRelationship Object

Represent a relationship between  **ModelTableColumn** objects. Used when programmatically creating relationships. Table 19 lists the properties of the **ModelRelationship** object.

 **Table 19. Properties of the ModelRelationship object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Active**|Read/Write| **Boolean**|When  **True**, the relationship is active.|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **ForeignKeyColumn**|Read-only| **ModelTableColumn**|Contains the  **ModelTableColumn** object that represents the foreign key column on the many side of the one-to-many relationship.|
| **ForeignKeyTable**|Read-only| **ModelTable**|Contains the  **ModelTable** object that represents the table on the many side of the one-to-many relationship.|
| **Parent**|Read-only| **Object**|Returns an  **Object** model object that represents the model the **ModelRelationship** object resides in.|
| **PrimaryKeyColumn**|Read-only| **ModelTableColumn**|Contains the  **ModelTableColumn** object that represents the primary key column in the table on the one side of the one-to-many relationship.|
| **PrimaryKeyTable**|Read-only| **ModelTable**|Contains the  **ModelTable** object that represents the table on the one side of the one-to-many relationship.|
 **ModelRelationship.Delete** Method

The  **Delete** method of the **ModelRelationship** object has no parameters. Deletes a relationship.


### ModelTables Collection

The  **ModelTables** collection contains a collection of **ModelTable** objects in the PPM. Table 20 lists the properties of the **ModelTables** collection.

 **Table 20. Properties of the ModelTables collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelTable** objects in the collection|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelTables** collection.|
 **ModelTables.Item** Method

Returns a single object from the  **ModelTables** collection. Table 21 lists the parameters of the Item method.

 **Table 21. Parameters of the ModelTables.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelTable Object

Represent a table in the  **Model** object. The **ModelTable** object is read only which means it cannot be created or edited through the object model. There is a **ModelTable** object for every table in the model. Table 22 lists the properties of the **ModelTable** object.

 **Table 22. Properties of the ModelTable object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **ModelTableColumns**|Read-only| **ModelTableColumns**|Collection of  **ModelTableColumn** objects that make up the **ModelTable** object.|
| **Name**|Read-only| **String**|Returns the name of the  **ModelTable** object.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the model the **ModelTable** object resides in.|
| **RecordCount**|Read-only| **Integer**|Returns the total row count for the  **ModelTable** object.|
| **SourceName**|Read-only| **String**|Name of table at the data source. If table has no data source (created in the model), the property will return an error.|
| **SourceWorkbookConnection**|Read-only| **WorkbookConnection**|Returns the workbook connection from which the  **ModelTable** object originated.|
 **ModelTable.Refresh** Method

The  **Refresh** method of the **ModelTable** object has no parameters. Refreshes the model table source connections.


### ModelTableColumns Collection

The  **ModelTableColumns** collection contains a collection of **ModelTableColumn** objects in the PPM. Table 23 lists the properties of the **ModelTableColumns** collection.

 **Table 23. Properties of the ModelTableColumns collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelTableColumn** objects in the collection|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelTableColumns** collection.|
 **ModelTableColumns.Item** Method

Returns a single object from the  **ModelTableColumns** collection. Table 24 lists the parameters of the **Item** method.

 **Table 24. Parameters of the ModelTableColumns.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelTableColumn Object

Represent a single column in the  **ModelTable** object. Used when programmatically creating relationships. Table 25 lists the properties of the **ModelTableColumn** object.

 **Table 25. Properties of the ModelTableColumn object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **DataType**|Read-only| **XlParameterDataType**|Returns the data type of the column.|
| **Name**|Read-only| **String**|Returns the name of the  **ModelTableColumn** object.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelTableColumn** object.|

### ModelTableNames Collection

The  **ModelTableNames** collection contains a collection of **ModelTableName** objects in the PPM. Table 26 lists the properties of the **ModelTableNames** collection.

 **Table 26. Properties of the ModelTableNames collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelTableName** objects in the collection|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelTableNames** object.|
 **ModelTableNames.Item** Method

Returns a single object from the  **ModelTableNames** collection. Table 27 lists the parameters of the **Item** method.

 **Table 27. Parameters of the ModelTableNames.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelTableNameChanges Collection

The  **ModelTableNameChanges** collection contains a collection of **ModelTableNameChange** objects in the PPM. Table 28 lists the properties of the **ModelTableNameChanges** collection.

 **Table 28. Properties of the ModelTableNameChanges collection**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Count**|Read-only| **Long**|Returns number of  **ModelTableNameChange** objects in the collection.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the parent object of the specified **ModelTableNameChanges** collection.|
 **ModelTableNameChanges.Item** Method

Returns a single object from the  **ModelTableNameChanges** collection. Table 29 lists the parameters of the **Item** method.

 **Table 29. Parameters of the ModelTableNameChanges.Item method**



|**Name**|**Required/Optional**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The index number or name of the object.|

### ModelTableNameChange Object

An object that represents the old and new name of a table which was renamed in the PPM. Table 30 lists the properties of the  **ModelTableNameChange** object.

 **Table 30. Properties of the ModelTableNameChange object**



|**Property**|**Read/Write**|**Type**|**Description**|
|:-----|:-----|:-----|:-----|
| **Application**|Read-only| **Application**|Returns an object that represents the Microsoft Excel application.|
| **Creator**|Read-only| **xlCreator**|Returns a 32-bit integer that indicates the application in which the specified object was created.|
| **Parent**|Read-only| **Object**|Returns an  **Object** that represents the model the **ModelTableNameChange** object resides in.|
| **TableNameNew**|Read-only| **String**|Returns the new name of the table.|
| **TableNameOld**|Read-only| **String**|Returns the old name of the table.|

## Conclusion
<a name="XLPowerPivotModel_Conclusion"> </a>

The PowerPivot add-in enables you to build your own cubes instead of using the default ones Excel creates for you behind Power tables. With this add-in, can see the cubes in a visual context and change cube-specific properties. The  **Model** object stores references to workbook connections and information about the Tables and Relationships contained within the PowerPivot Model.


## Additional Resources
<a name="XLPowerPivotModel_Additional"> </a>


-  [PowerPivot for Excel Tutorial Introduction](http://technet.microsoft.com/en-us/library/gg413497.aspx)
    
-  [PowerPivot for Excel Tutorial Sample Data](http://powerpivotsdr.codeplex.com/releases/view/35438)
    
-  [Using PowerPivot with Excel 2010](http://blogs.office.com/b/microsoft-excel/archive/2009/10/23/using-powerpivot-with-excel-2010.aspx)
    

