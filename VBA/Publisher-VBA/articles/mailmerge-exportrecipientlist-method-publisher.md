---
title: MailMerge.ExportRecipientList Method (Publisher)
keywords: vbapb10.chm6225941
f1_keywords:
- vbapb10.chm6225941
ms.prod: publisher
api_name:
- Publisher.MailMerge.ExportRecipientList
ms.assetid: 230d0f66-7368-51b7-8233-3fd54cfd0fe4
ms.date: 06/08/2017
---


# MailMerge.ExportRecipientList Method (Publisher)

Exports the list of mail-merge recipients to a Microsoft Office Access (.mdb) file or to a comma-delimited text (.csv) file.


## Syntax

 _expression_. **ExportRecipientList**( **_Filename_**,  **_FileType_**,  **_IncludedOnly_**)

 _expression_A variable that represents a  **MailMerge** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Required| **String**|The name of the file that will contain the list of recipients.|
|FileType|Optional| **PbRecipientListFileType**|The type of file to save. See Remarks for possible values.|
|IncludedOnly|Optional| **Boolean**|Whether to restrict entries in the recipient list to specific recipients.|

## Remarks

Possible values for the FileType parameter include the following constants from the  **PbRecipientListFileType** enumeration:



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbAsCsvFile**|1|Save as comma-delimited CSV file.|
| **pbAsMdbFile**|0|Save as Microsoft Office Access MDB file.|
The  **ExportRecipientList** method corresponds to the **Export recipient list to new file** command in the **E-mail Merge** and **Mail Merge** task panes in the Microsoft Publisher user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ExportRecipientList** method to export the list of mail-merge recipients to an Access database file. Before running this macro, ensure that the active document is connected to a data source. If the active document is not already connected to a data source, you can use the ** [MailMerge.OpenDataSource](mailmerge-opendatasource-method-publisher.md)** method to make the connection.

Also, before running the code, replace  _username_ in the folder path to the saved file with the name of a valid user on your computer, or replace the folder path and file name with a path and file name of your choice.

Note that the folder path used in this example is typical of folder paths in Microsoft Windows Vista. You must have permission to save files in the folder you designate.




```vb
Public Sub ExportRecipientList_Example() 
 
 Dim pubMailMerge As Publisher.MailMerge 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 pubMailMerge.ExportRecipientList "C:\Users\username\Documents\My Data Sources\MyAddressList", pbAsMdbFile, True 
 
End Sub
```


