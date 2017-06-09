---
title: MailMerge.OpenDataSource Method (Publisher)
keywords: vbapb10.chm6225937
f1_keywords:
- vbapb10.chm6225937
ms.prod: publisher
api_name:
- Publisher.MailMerge.OpenDataSource
ms.assetid: 4473e566-687f-595e-9fd6-a5483021cb48
ms.date: 06/08/2017
---


# MailMerge.OpenDataSource Method (Publisher)

Attaches a data source to the specified publication, which becomes a main publication if it is not one already.


## Syntax

 _expression_. **OpenDataSource**( **_bstrDataSource_**,  **_bstrConnect_**,  **_bstrTable_**,  **_fOpenExclusive_**,  **_fNeverPrompt_**)

 _expression_A variable that represents a  **MailMerge** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|bstrDataSource|Optional| **String**|The data source path and file name. You can specify a Microsoft Query (.qry) file instead of specifying a data source, a connection string, and a table name string; values in a Microsoft Query file override values for bstrConnect and bstrTable.|
|bstrConnect|Optional| **String**|A connection string.|
|bstrTable|Optional| **String**|The name of the table in the data source.|
|fOpenExclusive|Optional| **Long**| **True** to deny others access to the database. **False** allows others read/write permission to the database. The default value is **False**.|
|fNeverPrompt|Optional| **Long**| **Long**.  **True** never prompts when opening the data source. **False** displays the Data Link Properties dialog box. The default value is **False**.|

## Remarks




 **Note**  If you are using a data source for mail merge, you must add a catalog merge area to the publication page before you attach to the data source.


## Example

This example attaches a table from a database and denies everyone else write access to the database while it is opened. 

For this example to run properly, you must replace  _PathToFile_ with a valid file path and _TableName_ with a valid data source table name.




```vb
Sub AttachDataSource() 
 
    ActiveDocument.MailMerge.OpenDataSource _ 
        bstrDataSource:="PathToFile",  _ 
        bstrTable:="TableName", _ 
        fNeverPrompt:=True, fOpenExclusive:=True 
 
End Sub
```


