---
title: DataRecordsets Object (Visio)
keywords: vis_sdr.chm61000
f1_keywords:
- vis_sdr.chm61000
ms.prod: visio
api_name:
- Visio.DataRecordsets
ms.assetid: edf6d0dc-2f16-eee0-fd4c-ec4c9409179e
ms.date: 06/08/2017
---


# DataRecordsets Object (Visio)

The collection of  **DataRecordset** objects associated with a **Document** object.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the  **DataRecordsets** collection is **[Item](http://msdn.microsoft.com/library/8a289fb1-8cc5-eb76-efb1-c01f73c6340a%28Office.15%29.aspx)**.

Every Visio  **Document** object has a **DataRecordsets** collection, which is empty until you import data into Visio. To connect a Visio document to a data source, you add a **DataRecordset** object to the **DataRecordsets** collection of the document.

To add a  **DataRecordset** object to the **DataRecordsets** collection, you can use one of the following three methods, depending on the type of data source you want to connect to (OLEDB/ODBC or XML) and how you want to pass connection string and query command strings to Visio. By using the




-  **[DataRecordsets.Add](http://msdn.microsoft.com/library/9eb136ce-d543-75c3-3a72-cb23dfc8df78%28Office.15%29.aspx)** method, you can connect to an OLEDB or ODBC data source and pass connection and query command string information to Visio directly as method parameters.
    
-  **[DataRecordsets.AddFromConnectionFile](http://msdn.microsoft.com/library/7118bd4d-484b-dc22-e6f8-925376a5a67a%28Office.15%29.aspx)** method, you can connect to an OLEBD or ODBC data source by passing the method an Office Data Connection (ODC) file that contains the connection and query command string information you want to supply to Visio.
    
-  **[DataRecordsets.AddFromXML](http://msdn.microsoft.com/library/b75d7ecc-98d2-ae9b-608f-a9ec2b736ea6%28Office.15%29.aspx)** method, you pass the method an ADO classic XML string that contains all the data that you want to include in the data recordset.
    


Once you have created a data recordset, the connection string and query command string associated with the data recordset are represented by the  **[DataConnection.ConnectionString](http://msdn.microsoft.com/library/a1a6105f-64ee-1e0c-3b54-9831aec06bf4%28Office.15%29.aspx)** and **[DataRecordset.CommandString](http://msdn.microsoft.com/library/7d9151b0-db8c-a8ce-edea-7ef25d241e98%28Office.15%29.aspx)** properties respectively.


## Events



|**Name**|
|:-----|
|[BeforeDataRecordsetDelete](http://msdn.microsoft.com/library/6cb35848-51fe-653d-6cb3-a91e324bc6f3%28Office.15%29.aspx)|
|[DataRecordsetChanged](http://msdn.microsoft.com/library/44ee69e9-1c10-0d44-ccf4-d1787a261759%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/9f3fa9b0-2ca9-cf28-fa27-18eef4be179d%28Office.15%29.aspx)|
|[GetAllRefreshConflicts](http://msdn.microsoft.com/library/96d1c866-6c0d-f750-46a8-8257340ebd71%28Office.15%29.aspx)|
|[GetDataRowIDs](http://msdn.microsoft.com/library/d76874eb-c25b-df65-5d00-64de288d086e%28Office.15%29.aspx)|
|[GetMatchingRowsForRefreshConflict](http://msdn.microsoft.com/library/07526278-19db-ccbc-6785-095c73128879%28Office.15%29.aspx)|
|[GetPrimaryKey](http://msdn.microsoft.com/library/4f056424-4668-7859-5ed1-bd28a051ddc0%28Office.15%29.aspx)|
|[GetRowData](http://msdn.microsoft.com/library/969d7702-e78c-736f-87d8-c8e7e8c5a778%28Office.15%29.aspx)|
|[Refresh](http://msdn.microsoft.com/library/0a871f32-f24e-07c0-3cc6-a76f2a4ba2e2%28Office.15%29.aspx)|
|[RefreshUsingXML](http://msdn.microsoft.com/library/345935ab-b269-61dd-9ebe-e1f87b89bb11%28Office.15%29.aspx)|
|[RemoveRefreshConflict](http://msdn.microsoft.com/library/a92abdb7-f47c-b843-cacf-6acca68d9c66%28Office.15%29.aspx)|
|[SetPrimaryKey](http://msdn.microsoft.com/library/5ec125ff-b4a8-abcb-0d9d-140e97de6db2%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c602b9de-09b0-ca9b-a59b-4572be032a54%28Office.15%29.aspx)|
|[CommandString](http://msdn.microsoft.com/library/7d9151b0-db8c-a8ce-edea-7ef25d241e98%28Office.15%29.aspx)|
|[DataAsXML](http://msdn.microsoft.com/library/500dda1a-0747-57d0-f847-e3e1f72e96a3%28Office.15%29.aspx)|
|[DataColumns](http://msdn.microsoft.com/library/d22c07b9-3c92-fed4-72ed-6676ea64f1bf%28Office.15%29.aspx)|
|[DataConnection](http://msdn.microsoft.com/library/3425e9c4-4cd6-7553-2dbf-5e14b8a9a68a%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/ad59effe-9717-faa5-d427-0c22b693b626%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/419cdd3d-cb12-cbb6-5e47-d343b1a84d74%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/31d3b05b-31f7-538e-cff7-b4e62cb29187%28Office.15%29.aspx)|
|[LinkReplaceBehavior](http://msdn.microsoft.com/library/a49a9a44-1067-dfc6-0fb0-aee15064078b%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/6201d472-63ee-ac51-8d08-1bf1039d8b6d%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/0040cf96-c0b7-3f36-a7d6-76510ac5cab6%28Office.15%29.aspx)|
|[RefreshInterval](http://msdn.microsoft.com/library/3d108e6e-65af-05ea-77d2-a19d96f82c1e%28Office.15%29.aspx)|
|[RefreshSettings](http://msdn.microsoft.com/library/7647676c-0291-8c57-10d6-ca55fcee2bf5%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/b3df4d5a-bc10-db7f-3560-43519a7dae83%28Office.15%29.aspx)|
|[TimeRefreshed](http://msdn.microsoft.com/library/ebdf1acd-81f9-bd5e-48ba-d34100a8f702%28Office.15%29.aspx)|

