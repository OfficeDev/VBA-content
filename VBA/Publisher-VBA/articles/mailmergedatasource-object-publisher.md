---
title: MailMergeDataSource Object (Publisher)
keywords: vbapb10.chm6356991
f1_keywords:
- vbapb10.chm6356991
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource
ms.assetid: a02eb4fb-7db7-e533-c3ca-95bc4ca68e82
ms.date: 06/08/2017
---


# MailMergeDataSource Object (Publisher)

Represents the data source in a mail merge or catalog merge operation.
 


## Example

Use the  **[DataSource](mailmerge-datasource-property-publisher.md)** property to return the **MailMergeDataSource** object. The following example displays the name of the data source associated with the active publication.
 

 

```
Sub ShowDataSourceName() 
 If ActiveDocument.MailMerge.DataSource.Name <> "" Then _ 
 MsgBox ActiveDocument.MailMerge.DataSource.Name 
End Sub
```

The following example tests the open data source associated with the active publication to determine whether the LastName field includes the name Fuller.
 

 



```
Sub FindSelectedRecord() 
 With ActiveDocument.MailMerge 
 If .DataSource.FindRecord(FindText:="Fuller", _ 
 Field:="LastName") = True Then 
 MsgBox "Data was found" 
 End If 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ApplyFilter](mailmergedatasource-applyfilter-method-publisher.md)|
|[Close](mailmergedatasource-close-method-publisher.md)|
|[EditRecord](mailmergedatasource-editrecord-method-publisher.md)|
|[FindRecord](mailmergedatasource-findrecord-method-publisher.md)|
|[OpenRecipientsDialog](mailmergedatasource-openrecipientsdialog-method-publisher.md)|
|[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-publisher.md)|
|[SetAllIncludedFlags](mailmergedatasource-setallincludedflags-method-publisher.md)|
|[SetSortOrder](mailmergedatasource-setsortorder-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[ActiveRecord](mailmergedatasource-activerecord-property-publisher.md)|
|[Application](mailmergedatasource-application-property-publisher.md)|
|[ConnectString](mailmergedatasource-connectstring-property-publisher.md)|
|[DataFields](mailmergedatasource-datafields-property-publisher.md)|
|[DataSources](mailmergedatasource-datasources-property-publisher.md)|
|[EverValidated](mailmergedatasource-evervalidated-property-publisher.md)|
|[Filters](mailmergedatasource-filters-property-publisher.md)|
|[FirstRecord](mailmergedatasource-firstrecord-property-publisher.md)|
|[Included](mailmergedatasource-included-property-publisher.md)|
|[InvalidAddress](mailmergedatasource-invalidaddress-property-publisher.md)|
|[InvalidComments](mailmergedatasource-invalidcomments-property-publisher.md)|
|[IsMaster](mailmergedatasource-ismaster-property-publisher.md)|
|[LastRecord](mailmergedatasource-lastrecord-property-publisher.md)|
|[MappedDataFields](mailmergedatasource-mappeddatafields-property-publisher.md)|
|[Name](mailmergedatasource-name-property-publisher.md)|
|[Parent](mailmergedatasource-parent-property-publisher.md)|
|[RecordCount](mailmergedatasource-recordcount-property-publisher.md)|
|[TableName](mailmergedatasource-tablename-property-publisher.md)|
|[Type](mailmergedatasource-type-property-publisher.md)|
|[ValidatedClean](mailmergedatasource-validatedclean-property-publisher.md)|

