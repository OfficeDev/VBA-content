---
title: MailMerge.Execute Method (Publisher)
keywords: vbapb10.chm6225940
f1_keywords:
- vbapb10.chm6225940
ms.prod: publisher
api_name:
- Publisher.MailMerge.Execute
ms.assetid: edcabcc5-f2ce-53ce-d422-0d6fcb5f8a33
ms.date: 06/08/2017
---


# MailMerge.Execute Method (Publisher)

Performs the specified mail merge or catalog merge operation. Returns a  **[Document](document-object-publisher.md)** object that represents the new or existing publication specified as the destination of the merge results. Returns **Nothing** if the merge is executed to a printer.


## Syntax

 _expression_. **Execute**( **_Pause_**,  **_Destination_**,  **_Filename_**)

 _expression_A variable that represents a  **MailMerge** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Pause|Required| **Boolean**| **True** to have Microsoft Publisher pause and display a troubleshooting dialog box if a merge error is found. **False** to ignore errors during mail merge or catalog merge.|
|Destination|Optional| **PbMailMergeDestination**|The destination of the mail merge or catalog merge results. Specifying  **pbSendToPrinter** for a catalog merge results in a run-time error.|
|Filename|Optional| **String**|The file name of the publication to which you want to append the catalog merge results.|

### Return Value

Document


## Remarks

Destination can be one of these  **PbMailMergeDestination** constants. The default is **pbSendToPrinter**.



| **pbSendToPrinter**|
| **pbMergeToNewPublication**|
| **pbMergeToExistingPublication**|

## Example

This example executes a mail merge if the active publication is a main document with an attached data source.


```vb
Sub ExecuteMerge() 
 Dim mrgDocument As MailMerge 
 Set mrgDocument = ActiveDocument.MailMerge 
 If mrgDocument.DataSource.ConnectString <> "" Then 
 mrgDocument.Execute Pause:=False 
 End If 
End Sub
```


