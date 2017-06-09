---
title: About Linking Shapes to Data
ms.prod: visio
ms.assetid: 4664d0d8-ff3e-68d5-6c43-0019a53d63fd
ms.date: 06/08/2017
---


# About Linking Shapes to Data

 **Note**  Data-connectivity features are available only to licensed users of Microsoft Visio Professional 2013.

There are four aspects of data connectivity in Visio:

- Connecting to a data source
    
- Linking shapes to data
    
- Displaying linked data graphically
    
- Refreshing linked data that has changed in the data source, updating linked shapes, and resolving any subsequent conflicts that may arise
    
Typically, you approach these aspects in the order in which they are listed; that is, you first connect your Visio drawing to a data source, then link shapes in your drawing to data in the data source, display the data in linked shapes graphically, and refresh the linked data when necessary. 
Each of these aspects has objects and members associated with it in the Visio object model. This topic deals with the second of these aspects, linking shapes in your Visio drawing to data. For more information about the other aspects of data connectivity, see the following topics: 

-  [About Connecting to Data in Visio](about-connecting-to-data-in-visio.md)
    
-  [About Displaying Data Graphically](about-displaying-data-graphically-visio.md)
    
To connect your Visio drawing to a data source programmatically, you can use the Visio API for data connectivity, which includes the following objects and their associated members:

-  **[DataRecordsets](datarecordsets-object-visio.md)** collection
    
-  **[DataRecordset](datarecordset-object-visio.md)** object
    
-  **[DataConnection](dataconnection-object-visio.md)** object
    
-  **[DataRecordsetChangedEvent](datarecordsetchangedevent-object-visio.md)** object
    
-  **[DataColumns](datacolumns-object-visio.md)** collection
    
-  **[DataColumn](datacolumn-object-visio.md)** object
    
After you  [connect your Visio drawing to an external data source](about-connecting-to-data-in-visio.md), you can link the shapes in the drawing to data from that source programmatically. You can link one or more shapes to a single row of data in a data recordset or to multiple rows of data in different data recordsets. However, you cannot link shapes to multiple rows of data in the same recordset.
You can link existing shapes to data, one shape at a time or as a group; or, you can create shapes and link them to data simultaneously. You can specify the correspondence between shapes and data rows, if you know it, or you can let Visio determine the correspondence automatically, based on a comparison between existing shape data and data in the data recordset.
After you link shapes to data, you can display that data graphically by adding data graphics to shapes. For more information about data graphics, see  [About Displaying Data Graphically](about-displaying-data-graphically-visio.md).
The  **[DataRecordset](datarecordset-object-visio.md)** and **[DataColumn](datacolumn-object-visio.md)** objects and the **[DataColumns](datacolumns-object-visio.md)** collection expose several properties, methods, and events that facilitate data linking. In addition, several members of other objects in the Visio object model, including the **Application**,  **Document**,  **Page**,  **Selection**,  **Shape**, and  **Window** objects, are related to data-linking.

## Data-linking and Shape Data

Linking shapes to data relies on the fact that you can assign shape data to all Visio shapes. In versions of Visio earlier than Visio 2007, shape data were called custom properties.

To access and assign shape data in the Visio UI, right-click a shape, point to  **Data**, and then click  **Shape Data**. Alternatively, you can access and assign shape data manually or programmatically in the Visio ShapeSheet spreadsheet. To display the ShapeSheet spreadsheet (ShapeSheet) for a selected shape, right-click the shape and click  **Show ShapeSheet**. To see this command, you must be running Visio in developer mode. To run Visio in developer mode, click the  **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, select  **Run in developer mode**.

Within the ShapeSheet, shape data is contained in the Shape Data section (previously called the Custom Properties section). To maintain backwards-compatibility, existing object members retain "custom property" or "custom properties" in their name. If you do not assign shape data for a given shape, no Shape Data section appears in the ShapeSheet. You can add a Shape Data section to a ShapeSheet by displaying the ShapeSheet as described previously, right-clicking anywhere in the ShapeSheet window and clicking  **Insert Section**, selecting  **Shape Data**, and then clicking  **OK**.

After you link shapes to data, many of the columns of the Shape Data section correspond closely to the properties of the  **DataColumn** object. For example, the Label column in the Shape Data section, which provides the label that appears for a particular shape data item in the **Shape Data** dialog box, corresponds to the **[DataColumn.DisplayName](datacolumn-displayname-property-visio.md)** property, which controls the name that appears for the associated data column in the **External Data** window. For more information about working with the **DataColumn** object, see [Getting and Setting Data-column Properties](#getsetprops). 


## Identifying Shapes, Data Recordsets, and Data Rows

Visio uses unique ID numbers to identify shapes, recordsets, and data rows. Shape IDs are unique only within the scope of the page they are on. After you determine these numbers, you can pass them to methods of the Visio data-related objects to specify exactly how the shapes in your diagram should link to data rows in the available data recordsets.

To determine the ID for a shape, get the  **[Shape.ID](shape-id-property-visio.md)** property value. In addition, Visio also gives shapes unique IDs or GUIDs. The **[Page.ShapeIDsToUniqueIDs](page-shapeidstouniqueids-method-visio.md)** method takes an array of shape IDs, as well as an enumeration value from **[VisUniqueIDArgs](visuniqueidargs-enumeration-visio.md)** specifying whether to get, get or make, or delete shape GUIDs. The **Page.ShapeIDsToUniqueIDs** method also returns an array of unique IDs for the shapes passed in. Conversely, if you know the unique IDs of a set of shapes, you can use the **[Page.UniqueIDsToShapeIDs](page-uniqueidstoshapeids-method-visio.md)** method to obtain the shape IDs for those shapes. For a selection of shapes, use the **[Selection.GetIDs](selection-getids-method-visio.md)** method to get the shape IDs of the shapes.

To determine the ID for a  **DataRecordset** object you add to the **[DataRecordsets](datarecordsets-object-visio.md)** collection, get the **[DataRecordset.ID](datarecordset-id-property-visio.md)** property value. To determine the IDs for each of the rows in a data recordset, call the **[DataRecordset.GetDataRowIDs](datarecordset-getdatarowids-method-visio.md)** method, which returns an array of row IDs. For more information, see the section "Accessing Data in Data Recordsets Programmatically," in [About Connecting to Data in Visio](about-connecting-to-data-in-visio.md).


## Creating Shapes Linked to Data

When you want to create shapes, already linked to data, on a drawing page that either does not contain any shapes or contains shapes other than the ones you want to link, you can use the  **[Page.DropLinked](page-droplinked-method-visio.md)** and **[Page.DropManyLinkedU](page-dropmanylinkedu-method-visio.md)** methods to create one or more additional shapes already linked to data. These methods resemble the **Page.Drop** and **Page.DropManyU** methods in that they create additional shapes at a specified location on the page; but in addition, they create links between the new shapes and specified data rows in a specified data recordset.

The  **DropLinked** method returns the new, linked **Shape** object and takes the following parameters:


- ObjectToDrop The particular shape (a Rectangle shape, for example) you want to create.
    
- x The _x_-coordinate of the center of the new shape on the page.
    
- y The _y_-coordinate of the center of the new shape on the page.
    
- DataRecordsetID The value of the **ID** property of the **DataRecordset** object that contains the data row to link to.
    
- DataRowID The value of the **ID** property of the data row to link to.
    
- ApplyDataGraphicAfterLink A **Boolean** value specifying whether to apply the shape's data graphic automatically if it already has one, or if not, whether to apply the most recently used data graphic. The default is not to apply a data graphic. For more information on data graphics, see [About Displaying Data Graphically](about-displaying-data-graphically-visio.md). 
    
The following sample code shows how to use the  **DropLinked** method to create a shape on the active drawing page, centered at page coordinates (2, 2), and linked to a data row. It takes the **DataRecordset** object passed in, gets its ID, and then passes that ID, along with ID of the data row to link to, to the **DropLinked** method. The dropped shape is a simple rectangle from the Basic_U.VSS stencil, which the code opens, docked in the Visio drawing window.

In this example, the ID of the data row is set to 1; before running the code, ensure that a row with that ID exists, or change the ID value in the code.




```vb
Public Sub DropLinkedShape(vsoDataRecordset As Visio.DataRecordset) 
 
    Dim vsoShape As Visio.Shape 
    Dim vsoMaster As Visio.Master 
    Dim dblX As Double 
    Dim dblY As Double  
    Dim lngRowID As Long 
    Dim lngDataRecordsetID As Long 
 
    lngDataRecordsetID = vsoDataRecordset.ID 
    Set vsoMaster = Visio.Documents.OpenEx("Basic_U.VSS", 0).Masters("Rectangle") 
    x = 2 
    y = 2 
    lngRowID = 1 
    Set vsoShape = ActivePage.DropLinked(vsoMaster, dblX, dblY, lngDataRecordsetID, lngRowID, True) 
 
End Sub
```

The  **DropManyLinkedU** method similarly creates a set of linked shapes, returned as an array of shape IDs. It takes as parameters arrays of shapes to drop, coordinates, and data rows to link to. Entries at corresponding array-index positions determine how shapes and data rows are related and where on the page individual shapes are dropped.


## Linking Existing Shapes to Data

When you know exactly how one or more existing shapes in a Visio drawing and one or more rows in a data recordset correspond to one another, you can link the existing shapes to data in the following ways:


- Link a single shape to a single data row
    
- Link a selection of shapes to one or more data rows
    
- Link multiple shapes to multiple data rows
    
In addition, if you do not know the exact shape to data mapping, you can direct Visio to make the best match possible, based on limited matching information that you provide.


## Linking a Single Shape to a Data Row

To link a single shape to a single data row, use the  **[Shape.LinkToData](shape-linktodata-method-visio.md)** method. This method takes a data recordset ID and data row ID as well as an optional **Boolean** flag specifying whether to display the linked data in a data graphic. The default is to display the data graphic.


## Linking Multiple Shapes to Data

Two members of the  **Selection** object, the **[Selection.LinkToData](selection-linktodata-method-visio.md)** and **[Selection.AutomaticLink](selection-automaticlink-method-visio.md)** methods, as well as the **[Page.LinkShapesToDataRows](page-linkshapestodatarows-method-visio.md)** method, make it possible to link one or more existing shapes in a selection to data.

The  **[Selection.LinkToData](selection-linktodata-method-visio.md)** method functions much like the same method of the **Shape** object, except that it links a selection of shapes, instead of a single shape, to a single data row.

If you are unsure about the correspondence between shapes and data rows, but know a match exists between a specific attribute of every shape and the data in one column in the data recordset, the  **[Selection.AutomaticLink](selection-automaticlink-method-visio.md)** method provides a means to link a selection of existing shapes to multiple rows of data. Note that it must be the same attribute for all shapes. For more information about this method, see [Linking to Data Automatically](#linktodataauto).

The  **[Page.LinkShapesToDataRows](page-linkshapestodatarows-method-visio.md)** method is similar to the **Selection.LinkToData** method in that it links multiple shapes. However, you use this method to link shapes on the same page, rather than shapes in a selection, to data. The **LinkShapesToDataRows** method links shapes to multiple data rows, whereas the **LinkToData** method links multiple shapes to a single row. To link shapes, pass the **LinkShapesToDataRows** method a pair of arrays: one for shapes, and one for data rows. Note that the matching array positions must correspond. As a result, for example, the shape at position 1 in the shape array is linked to the data at position 1 in the data row array. Once again, when you call the method, you can optionally specify whether to apply an existing data graphic to linked shapes.


## Linking to Data Automatically

You can use the  **Selection.AutomaticLink** method to link shape data values in selected shapes—that is, shapes assigned to a Selection object—to data rows in a data recordset automatically—that is, without specifying the exact correspondence of all shapes and data rows. To provide Visio with enough information to create the links, however, you must supply at least one set of matching data: the name of a column in the database, a shape attribute type, and, if necessary, a shape value, all at the same index position of the corresponding arrays you pass to the method.

The shape attribute type indicates the attribute of the shape to base the matching upon. The attribute can be the value of a shape data item (formerly known as a custom property value), shape text, or another of the values specified in the  **[VisAutoLinkFieldTypes](visautolinkfieldtypes-enumeration-visio.md)** enumeration.


 **Note**  For example, say that your drawing contains a selection of shapes representing different employees. Their shape text identifies the shapes, which in this case would be the respective employee's names. (You could use some of the employee names from the OrgData.xls workbook that ships with Visio, and then connect to that data source. By default, OrgData.xls is installed at the following path: C:\Program Files\Microsoft Office\Office15\Visio Content\[ _langID_], where  _langID_ varies by country or region.) On some computers, the path might include "Program Files (x86)" instead of "Program Files."

To connect these shapes to a database where each employee's data constitutes a row in the database, you pass the following parameters to the  **AutomaticLink** method:


- DataRecordsetID The value of the **ID** property of the **DataRecordset** object that contains the data rows to link to. In the example that follows, we pass an existing data recordset to the procedure and get its ID.
    
- ColumnNames() A string array consisting of names of columns in the database. At least one position in the array must have a value that corresponds to the values in the same position in theAutoLinkFieldTypes andFieldNames arrays. In the following example, we pass an array that contains the "Name" column name at array position 0.
    
- AutoLinkFieldTypes() An array of **Long** values from the **VisAutoLinkFieldTypes** enumeration, consisting of shape attribute types. At least one position in the array must have a value that corresponds to the values in the same position in theColumnNames andFieldNames arrays. In the following example, we pass the enumeration value **visAutoLinkShapeText** at array position 0.
    
- FieldNames() A string array consisting of shape values. At least one position in the FieldNames array must have a value that corresponds to the values in the same position in theColumnNames andAutoLinkFieldTypes arrays.
    
- For most values of  _AutoLinkFieldTypes_, for example,  **visAutoLinkShapeText**, it is not necessary to specify the  _FieldNames_ value; you can pass the null value instead. That is the case in our example, so we pass an empty string. However, when you pass the **visAutoLinkCustPropsLabel**,  **visAutoLinkUserRowName**,  **visAutoLinkPropRowNameU**, or  **visAutoLinkUserRowNameU** values of _AutoLinkFieldTypes_, you must pass a value for  _FieldNames_ to fully specify the shape data item to compare to the data column name.
    
- AutoLinkBehavior A value from the **VisAutoLinkBehaviors** enumeration. These enumerated values provide options to customize the method, for example, to replace existing links with new ones. The following example passes the default value, 0.
    
- ShapeIDs() An array that the method fills with the IDs of linked shapes when it returns.
    
The following sample shows one way to use the  **AutomaticLink** method to link shapes and data automatically. The sample assumes that you have connected your drawing to data in the OrgData.xls sample workbook, as explained above. Note that the code requires that the first column of data be named "Name," as is the case in OrgData.xls. Note also that the shape text of each of the shapes in your drawing that you want to link to data must match one of the names in the "Name" column in OrgData.xls.




```vb
Public Sub LinkToDataAutomatically(vsoDataRecordset As Visio.DataRecordset) 
 
    Dim vsoSelection As Visio.Selection 
    Dim columnNames(1) As String 
    Dim fieldTypes(1) As Long 
    Dim fieldNames(1) As String 
    Dim shapesLinked() As Long 
 
    columnNames(0) = "Name" 
    fieldTypes(0) = Visio.VisAutoLinkFieldTypes.visAutoLinkShapeText 
    fieldNames(0) = "" 
    ActiveWindow.DeselectAll 
    ActiveWindow.SelectAll 
    Set vsoSelection = ActiveWindow.Selection 
    vsoSelection.AutomaticLink vsoDataRecordset.ID, _ 
                    columnNames, _ 
                    fieldTypes, _ 
                    fieldNames, 0, shapesLinked 
 
End Sub
```


## Discovering Links between Shapes and Data

Use the following methods to determine which shapes are linked to data. Knowing how shapes are linked to data can help prevent conflicts and broken links:


-  **[Page.GetShapesLinkedToData](page-getshapeslinkedtodata-method-visio.md)**
    
-  **[Page.GetShapesLinkedToDataRow](page-getshapeslinkedtodatarow-method-visio.md)**
    
-  **[Shape.GetLinkedDataRow](shape-getlinkeddatarow-method-visio.md)**
    
-  **[Shape.GetCustomPropertyLinkedColumn](shape-getcustompropertylinkedcolumn-method-visio.md)**
    
-  **[Shape.GetCustomPropertiesLinkedToData](shape-getcustompropertieslinkedtodata-method-visio.md)**
    
-  **[Shape.IsCustomPropertyLinked](shape-iscustompropertylinked-method-visio.md)**
    

## Breaking Links between Shapes and Data

As their names imply, you can use the  **[Shape.BreakLinkToData](shape-breaklinktodata-method-visio.md)** and **[Selection.BreakLinkToData](selection-breaklinktodata-method-visio.md)** methods to break existing links between shapes and data programmatically. In addition, various changes made in the UI can break these links. For example, when users delete a data recordset, linked row, or linked shape, or when users click **Unlink from Row** on a shape's shortcut menu or Unlink on a row's shortcut menu, they can cause broken links.

Except when a user deletes a data recordset, row, or shape from the UI, all of these actions fire the  **[Shape.ShapeLinkDeleted](shape-shapelinkdeleted-event-visio.md)** event. You can also use the methods listed in the previous section to determine link status.


## Getting and Setting Data-column Properties

Every  **DataRecordset** object contains a **DataColumns** collection of all the **DataColumn** objects associated with the **DataRecordset** object. These objects allow you to map data columns to cells in the Shape Data section of the ShapeSheet.

The following sample shows how to get the value of the Label cell in the Shape Data section for the first column in the data recordset passed to the method and display it in the  **Immediate** window. Then it sets the value and displays the new value.

Changing this value changes the label of the shape data item in the  **Shape Data** dialog box for all shapes linked to rows in the data recordset. To get and set the Label cell value, we pass the **visDataColumnPropertyDisplayName** value from the **[VisDataColumnProperties](visdatacolumnproperties-enumeration-visio.md)** enumeration to the **[DataColumn.GetProperty](datacolumn-getproperty-method-visio.md)** and **[DataColumn.SetProperty](datacolumn-setproperty-method-visio.md)** methods.




```vb
Public Sub ChangeColumnProperties(vsoDataRecordset As Visio.DataRecordset) 
 
    Dim strPropertyName As String 
    Dim strNewName As String 
    Dim vsoDataColumn As Visio.DataColumn 
 
    strNewName = "New Property Name" 
    Set vsoDataColumn = vsoDataRecordset.DataColumns(1) 
    strPropertyName = vsoDataColumn.GetProperty(visDataColumnPropertyDisplayName) 
    Debug.Print strPropertyName 
 
    vsoDataColumn.SetProperty visDataColumnPropertyDisplayName, strNewName 
    strPropertyName = vsoDataColumn.GetProperty(visDataColumnPropertyDisplayName) 
    Debug.Print strPropertyName 
 
End Sub
```


## Refreshing Linked Data and Resolving Conflicts

When data changes in the data source to which your drawing is connected, you can refresh the data in your Visio drawing to reflect those changes. You can specify that Visio refresh data automatically at a specified interval by setting the  **[DataRecordset.RefreshInterval](datarecordset-refreshinterval-property-visio.md)** property. You can refresh data programmatically by calling the **[DataRecordset.Refresh](datarecordset-refresh-method-visio.md)** method.

In addition, you can resolve any conflicts in the relationship between shapes and rows of data. For example, conflicts can occur when you refresh the data recordset and some data rows to which shapes were linked before the refresh operation no longer exist, because of changes to the data source. Other conflicts are possible when two or more rows in the refreshed recordset have identical primary keys.


## Refreshing Linked Data Automatically

When you create a  **DataRecordset** object, its **RefreshInterval** property value is set to the default, 0. This setting indicates that data is not refreshed automatically. By setting **DataRecordset.RefreshInterval** to a positive **Long** value, you can specify the time in minutes between automatic refreshes. The minimum interval you can specify is one minute. This setting corresponds to the value a user can set in the **Configure Refresh** dialog box.

To determine the date and time of the last refresh operation, get the  **[DataRecordset.TimeRefreshed](datarecordset-timerefreshed-property-visio.md)** property.

Additionally, the  **[DataRecordset.RefreshSettings](datarecordset-refreshsettings-property-visio.md)** property allows you to customize automatic refreshes of data. By setting this property to a combination of the values in the**[VisRefreshSettings](visrefreshsettings-enumeration-visio.md)** enumeration, you can specify that either or both of the following occur:


- The UI for reconciling refresh conflicts (the  **Refresh Conflicts** task pane) is disabled. (See the next section for more information.)
    
- Refresh operations automatically overwrite changes to data made in the UI. The default value for this property is 0, meaning that neither of these events occur.
    

## Identifying Data Recordset Rows for Refresh Operations

Because shapes are linked by their shape IDs to specific data rows, when Visio refreshes linked data, it must determine which rows in the linked data recordset or recordsets were added, changed, or removed since the last time the data was refreshed. To identify these rows, Visio uses the row IDs assigned to the rows in the data recordset. Visio can assign these row IDs two ways, depending on whether you designated primary keys for the data recordset when you created it.


## Refreshing Data Recordsets that Do Not Have Primary Keys

 When you create a data recordset, Visio assigns row IDs to all the rows in the recordset based on the existing order of the rows in the data source. Accordingly, the first row in the recordset is always assigned row ID 1, the second row ID 2 and so forth.

Subsequently, you can add or remove data rows from the original data source. Then, when you refresh the data, the data recordset reflects those changes. As a result, row order in the data recordset may change.

For example, in a five-row data recordset, if the fourth row in the data source is removed, when Visio refreshes the data recordset connected to that data source, the fifth row in the data recordset becomes the new fourth row and is assigned row ID 4. Row ID 5 is removed from the data recordset.

As a result, shapes linked to row ID 5 loose their links, and shapes linked to row ID 4 now get data from the row previously in the fifth position. As you can see, not assigning primary keys to data recordsets when you create them can result in broken links between shapes and data, or in Visio linking shapes to rows other than the ones to which you want them linked.


## Refreshing Data Recordsets that Have Primary Keys

You can help prevent these broken or mismatched links by assigning primary keys to data recordsets. A primary key identifies the name of the data column or columns that contain unique identifiers for each row. The value in the primary key column for each row uniquely identifies that row in the data recordset. Primary keys are often ID values, but you can set any column or combination of columns to be the primary key. However, to get consistent results when you refresh data, it is essential that you make the primary key column value (or set of values for multiple primary key columns) unique for each row.

As a result, when you refresh or when Visio refreshes a data recordset that includes primary keys, its rows retain the same row IDs they had before the refresh operation. Because Visio links shapes to data rows by ID—shape ID to row ID—and because row IDs remain the same after a refresh operation, data-linked shapes remain linked to the correct row. Note that row IDs are never recycled for a given a data recordset.

You can use the  **[DataRecordset.GetPrimaryKey](datarecordset-getprimarykey-method-visio.md)** method to determine the existing primary key for a data recordset, if one is specified. This method returns the primary key setting for the data recordset, as a value from the**[VisPrimaryKeySettings](visprimarykeysettings-enumeration-visio.md)** enumeration. You can use single or composite primary keys. A single key bases row identification on the values in a single column. A composite primary key uses two or more columns to identify a row uniquely.

If the primary key setting is  **visKeySingle** or **visKeyComposite**, the method also returns an array of primary key column-name strings. If the primary key setting is  **visKeyRowOrder**, the default, the method returns an empty array of primary keys.

Likewise, you can use the  **[DataRecordset.SetPrimaryKey](datarecordset-setprimarykey-method-visio.md)** method to specify the primary key setting for the data recordset, as well as the name of the column or columns that you want to set as the primary key column or columns. Once again, when you set primary keys, make sure that the column or columns you pick to be primary key columns contain unique values (or value sets) for each row.


## Refreshing Linked Data Programmatically

To refresh a connected data recordset programmatically, call the  **DataRecordset.Refresh** method.

Calling this method executes the query string associated with the data recordset and then updates the linked shapes with the data returned by the query. Calling the  **Refresh** method on a particular **DataRecordset** object results in refreshing all other **DataRecordset** objects associated with the same **[DataConnection](dataconnection-object-visio.md)** object (that is, having the same value for their **DataConnection** property). **DataRecordset** objects sharing the same **DataConnection** property value are called transacted data recordsets.

If calling  **Refresh** results in conflicts, Visio displays the **Refresh Conflicts** task pane in the UI, unless you set the **RefreshSettings** property to include the **visRefreshNoReconciliationUI** enumerated value.

Before you refresh linked data, if you want to change the query Visio uses to retrieve the data to query a different table in the same database, set the  **[DataRecordset.CommandString](datarecordset-commandstring-property-visio.md)** property to a new value. To connect to an entirely new data source, set both the **DataRecordset.CommandString** and **[DataConnection.ConnectionString](dataconnection-connectionstring-property-visio.md)** property values.

The  **[DataRecordset.GetLastDataError](datarecordsets-getlastdataerror-method-visio.md)** method gets the Active X Data Objects (ADO) error code, ADO description, and data recordset ID associated with the most recent error that resulted from adding a new data recordset or refreshing the data in an existing one.


## Identifying and Resolving Conflicts

When you or Visio refreshes data and a resulting conflict occurs, you can use the  **[DataRecordset.GetAllRefreshConflicts](datarecordset-getallrefreshconflicts-method-visio.md)** and **[DataRecordset.GetMatchingRowsForRefreshConflict](datarecordset-getmatchingrowsforrefreshconflict-method-visio.md)** methods to determine why the conflict arose. The **GetAllRefreshConflicts** method returns an array of shapes for which a conflict exists between data in the shape and data in the data-recordset row to which the shape is linked. To determine which data-recordset rows produced the conflict, you can then pass each of these shapes to the **GetMatchingRowsForRefreshConflict** method, which returns an array of rows that are in conflict.

Rows in the data recordset can conflict when two or more of them have identical primary keys, and may link to the same shape. When this occurs,  **GetMatchingRowsForRefreshConflict** returns an array containing at least two row IDs.

Conflicts can also occur when a previously data-linked row from the data recordset is removed. When this occurs, the method returns an empty array.

To remove the conflict, pass the shape to the  **[DataRecordset.RemoveRefreshConflict](datarecordset-removerefreshconflict-method-visio.md)** method, which removes the conflicting information from the current document.


