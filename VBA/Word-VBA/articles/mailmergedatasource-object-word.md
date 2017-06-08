---
title: MailMergeDataSource Object (Word)
keywords: vbawd10.chm2333
f1_keywords:
- vbawd10.chm2333
ms.prod: word
api_name:
- Word.MailMergeDataSource
ms.assetid: f86f7d3c-d7ab-45e8-21e7-fd5a426e0391
ms.date: 06/08/2017
---


# MailMergeDataSource Object (Word)

Represents the mail merge data source in a mail merge operation.


## Remarks

Use the  **DataSource** property to return the **MailMergeDataSource** object. The following example displays the name of the data source associated with the active document.


```
If ActiveDocument.MailMerge.DataSource.Name <> "" Then _ 
 MsgBox ActiveDocument.MailMerge.DataSource.Name
```

The following example displays the field names in the data source associated with the active document.




```
For Each aField In ActiveDocument.MailMerge.DataSource.FieldNames 
 MsgBox aField.Name 
Next aField
```

The following example opens the data source associated with Form letter.doc and determines whether the FirstName field includes the name "Kate."




```
With Documents("Form letter.doc").MailMerge 
 .EditDataSource 
 If .DataSource.FindRecord(FindText:="Kate", _ 
 Field:="FirstName") = True Then 
 MsgBox "Data was found" 
 End If 
End With
```


## Methods



|**Name**|
|:-----|
|[Close](mailmergedatasource-close-method-word.md)|
|[FindRecord](mailmergedatasource-findrecord-method-word.md)|
|[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-word.md)|
|[SetAllIncludedFlags](mailmergedatasource-setallincludedflags-method-word.md)|

## Properties



|**Name**|
|:-----|
|[ActiveRecord](mailmergedatasource-activerecord-property-word.md)|
|[Application](mailmergedatasource-application-property-word.md)|
|[ConnectString](mailmergedatasource-connectstring-property-word.md)|
|[Creator](mailmergedatasource-creator-property-word.md)|
|[DataFields](mailmergedatasource-datafields-property-word.md)|
|[FieldNames](mailmergedatasource-fieldnames-property-word.md)|
|[FirstRecord](mailmergedatasource-firstrecord-property-word.md)|
|[HeaderSourceName](mailmergedatasource-headersourcename-property-word.md)|
|[HeaderSourceType](mailmergedatasource-headersourcetype-property-word.md)|
|[Included](mailmergedatasource-included-property-word.md)|
|[InvalidAddress](mailmergedatasource-invalidaddress-property-word.md)|
|[InvalidComments](mailmergedatasource-invalidcomments-property-word.md)|
|[LastRecord](mailmergedatasource-lastrecord-property-word.md)|
|[MappedDataFields](mailmergedatasource-mappeddatafields-property-word.md)|
|[Name](mailmergedatasource-name-property-word.md)|
|[Parent](mailmergedatasource-parent-property-word.md)|
|[QueryString](mailmergedatasource-querystring-property-word.md)|
|[RecordCount](mailmergedatasource-recordcount-property-word.md)|
|[TableName](mailmergedatasource-tablename-property-word.md)|
|[Type](mailmergedatasource-type-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
