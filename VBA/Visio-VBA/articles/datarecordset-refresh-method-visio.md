---
title: DataRecordset.Refresh Method (Visio)
keywords: vis_sdr.chm16460320
f1_keywords:
- vis_sdr.chm16460320
ms.prod: visio
api_name:
- Visio.DataRecordset.Refresh
ms.assetid: 0a871f32-f24e-07c0-3cc6-a76f2a4ba2e2
ms.date: 06/08/2017
---


# DataRecordset.Refresh Method (Visio)

Executes the query string associated with the connected (non-XML-based)  **[DataRecordset](datarecordset-object-visio.md)** and updates linked shapes with new data from the data source returned by the query.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Refresh**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

Nothing


## Remarks

Calling the  **Refresh** method on a particular **DataRecordset** object results in refreshing all other **DataRecordset** objects associated with the same **[DataConnection](dataconnection-object-visio.md)** object (that is, having the same value for their **[DataConnection](datarecordset-dataconnection-property-visio.md)** property). **DataRecordset** objects sharing the same **DataConnection** property value are called _transacted_ data recordsets. The **Refresh** method must be called on a data recordset that is associated with a **DataConnection** objecct.

If you call  **Refresh** on a data recordset not associated with a **DataConnection** object (one that was created by using the **[DataRecordsets.AddFromXML](datarecordsets-addfromxml-method-visio.md)** method), the **Refresh** method will return an error.

If calling  **Refresh** results in conflicts, Visio displays the **Refresh Conflicts** task pane in the user interface, unless you set the **[DataRecordset.RefreshSettings](datarecordset-refreshsettings-property-visio.md)** property to include the **visRefreshNoReconciliationUI** enumerated value.

Before refreshing linked data, if you want to change the query string Visio uses to retrieve the data to query a different table in the same database, set the  **[DataRecordset.CommandString](datarecordset-commandstring-property-visio.md)** property to a new value. To connect to an entirely new data source, set both the **[DataRecordset.CommandString](datarecordset-commandstring-property-visio.md)** and **[DataConnection.ConnectionString](dataconnection-connectionstring-property-visio.md)** property values.

When you refresh data and a conflict occurs, you can use the  **[DataRecordset.GetAllRefreshConflicts](datarecordset-getallrefreshconflicts-method-visio.md)** and **[DataRecordset.GetMatchingRowsForRefreshConflict](datarecordset-getmatchingrowsforrefreshconflict-method-visio.md)** methods to determine why the conflict arose.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Refresh** method to refresh the data in a **DataRecordset** object from the **DataRecordsets** collection of the current document. It gets the count of all data recordsets associated with the current document and refreshes the one most recently added. It also refreshes any other data recordsets associated with the current document that share a common data connection with the one being refreshed.

Before you run this macro, make sure that the current document contains at least one data recordset, and that the most recently added data recordset is connected (non-XML-based).




```vb
Public Sub Refresh_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intCount As Integer 
 
    intCount = ThisDocument.DataRecordsets.Count 
    Set vsoDataRecordset = ThisDocument.DataRecordsets(intCount) 
    vsoDataRecordset.Refresh 
 
End Sub
```


