---
title: DataRecordset Object (Visio)
keywords: vis_sdr.chm61005
f1_keywords:
- vis_sdr.chm61005
ms.prod: visio
api_name:
- Visio.DataRecordset
ms.assetid: 272d5fbb-d8a7-1fe8-07a3-7d7f71b62936
ms.date: 06/08/2017
---


# DataRecordset Object (Visio)

Stores, formats, refreshes, and exposes data queried from a database in Microsoft Visio.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

Microsoft Visio can connect to data from a variety of sources, including the following:




- Microsoft Excel worksheets
    
- Microsoft Access databases
    
- Microsoft SQL Server databases
    
- Microsoft SharePoint lists
    
- Other OLEDB/ODBC data sources, such as Oracle databases
    
- XML files that adhere to the ADO classic XML schema
    


Every Visio  **Document** object has a **DataRecordsets** collection, which is empty until a connection is made to a data source. To connect a Visio document to a data source, you add a **DataRecordset** object to the **DataRecordsets** collection of the document.

To add a  **DataRecordset** object to the **DataRecordsets** collection, you can use one of the following three methods, depending on the type of data source you want to connect to (OLEDB/ODBC or XML) and how you want to pass connection string and query command strings to Visio. By using the




-  **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method, you can connect to an OLEDB or ODBC data source and pass connection and query command string information to Visio directly as method parameters.
    
-  **[DataRecordsets.AddFromConnectionFile](datarecordsets-addfromconnectionfile-method-visio.md)** method, you can connect to an OLEBD or ODBC data source by passing the method an Office Data Connection (ODC) file that contains the connection and query command string information you want to supply to Visio.
    
-  **[DataRecordsets.AddFromXML](datarecordsets-addfromxml-method-visio.md)** method, you pass the method an ADO classic XML string that contains all the data that you want to include in the data recordset.
    


Once you have created a data recordset, the connection string and query command string associated with the data recordset are represented by the  **[DataConnection.ConnectionString](dataconnection-connectionstring-property-visio.md)** and **[DataRecordset.CommandString](datarecordset-commandstring-property-visio.md)** properties respectively.

If the data recordset is associated with a  **[DataConnection](dataconnection-object-visio.md)** object—that is, if you added it to the **DataRecordsets** collection by using either the **Add** or **AddFromConnectionFile** method—you can use the **[DataConnection](datarecordset-dataconnection-property-visio.md)** property of the **DataRecordset** object to get the associated **DataConnection** object.

The default property of a  **DataRecordset** object is **[ID](datarecordset-id-property-visio.md)** . The **ID** property value identifies the **DataRecordset** in the **DataRecordsets** collection, and is unique within the collection for any given document.

You can use the  **[Name](datarecordset-name-property-visio.md)** property of the **DataRecordset** object to associate a display name with the data recordset.

You can use the  **[GetDataRowIDs](datarecordset-getdatarowids-method-visio.md)** method to get an array of the IDs of all the rows in a data recordset, where each row represents a single data record. Once you have retrieved the data-row IDs in this manner, you can use the **[GetRowData](datarecordset-getrowdata-method-visio.md)** method to get all the data stored in each column in the data row.

You can use the  **[DataColumns](datarecordset-datacolumns-property-visio.md)** property of the **DataRecordset** object to get the **[DataColumn](datacolumn-object-visio.md)** object associated with the data recordset. The **DataColumn** object exposes methods and properties that you can use to customize the mapping of data columns to cells in the Shape Data section of the Visio ShapeSheet spreadsheet for shapes linked to data.

Setting a primary key column for a data recordset can help prevent broken links between shapes and data when data is refreshed. You can get and set the primary key column by using the  **[GetPrimaryKey](datarecordset-getprimarykey-method-visio.md)** and **[SetPrimaryKey](datarecordset-setprimarykey-method-visio.md)** methods respectively.

When data changes in the data source, you can refresh the data in a connected (non-XML) data recordset to reflect those changes. You can specify that Visio refresh data automatically at a specified interval by setting the  **[RefreshInterval](datarecordset-refreshinterval-property-visio.md)** property, or you can refresh data programmatically by calling the **[Refresh](datarecordset-refresh-method-visio.md)** method.

When you refresh data from a data source that has changed since the last time you refreshed data, conflicts can occur. Conflicts can result when a single shape is linked to more than one row in the same data source, or when a shape is linked to a row in the data source that has been deleted. You can discover and resolve the conflicts that arise from refreshing data by using the  **[GetAllRefreshConflicts](datarecordset-getallrefreshconflicts-method-visio.md)** , **[GetMatchingRowsForRefreshConflict](datarecordset-getmatchingrowsforrefreshconflict-method-visio.md)** , and **[RemoveRefreshConflict](datarecordset-removerefreshconflict-method-visio.md)** methods.




 **Note**  When you save a Visio document that contains one or more data recordsets, all the data in the recordset is saved in the Visio file. For recordsets that contain a large amount of data, this can create large Visio files, which can affect performance. Consequently, you should consider filtering large data sources before importing them into Visio.


