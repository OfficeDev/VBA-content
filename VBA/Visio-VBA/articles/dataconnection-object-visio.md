---
title: DataConnection Object (Visio)
keywords: vis_sdr.chm61010
f1_keywords:
- vis_sdr.chm61010
ms.prod: visio
api_name:
- Visio.DataConnection
ms.assetid: db21a645-d24d-253f-11ee-c75261d0896b
ms.date: 06/08/2017
---


# DataConnection Object (Visio)

Abstracts communication between one or more  **[DataRecordset](datarecordset-object-visio.md)** objects and a non-XML data source.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the  **DataConnection** object is **ID** .

When you add a new  **DataRecordset** object to the **DataRecordsets** collection (by using a method other than **[DataRecordsets.AddFromXML](datarecordsets-addfromxml-method-visio.md)** ) and you do not specify an existing **DataConnection** object (by passing the connection string associated with it to the **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method), Visio creates a new **DataConnection** object.

The  **DataConnection** object exposes properties that make it possible to access data-source connection settings:




-  The **[ConnectionString](dataconnection-connectionstring-property-visio.md)** property gets or sets the connection string used to access an existing **DataConnection** object or to create a new **DataConnection** object. Note that setting this property to a new value does not immediately change the connection—Visio re-evaluates this property only when the **[DataRecordset.Refresh ](datarecordset-refresh-method-visio.md)** method is called.
    
- The  **[Timeout](dataconnection-timeout-property-visio.md)** property determines how long (in seconds) Visio should attempt to establish a data-source connection before terminating the connection attempt and generating an error. The default is 15 seconds.
    
- The  **[FileName](dataconnection-filename-property-visio.md)** property gets or sets the name of the Office Data Connection (ODC) file that contains the data-source connection and query information used to create a new connection and to refresh data from an existing connection.
    


Multiple  **DataRecordset** objects can share the same **DataConnection** object. When any of the data recordsets that share a data connection are refreshed, all are refreshed.


