---
title: Application.MailMergeAfterRecordMerge Event (Word)
keywords: vbawd10.chm4000017
f1_keywords:
- vbawd10.chm4000017
ms.prod: word
api_name:
- Word.Application.MailMergeAfterRecordMerge
ms.assetid: 6f483874-3999-815d-28b3-69fef89ed2be
ms.date: 06/08/2017
---


# Application.MailMergeAfterRecordMerge Event (Word)

Occurs after each record in the data source successfully merges in a mail merge.


## Syntax

 _expression_ . **Private Sub object_MailMergeAfterRecordMerge**( **_ByVal Doc As Document_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|

## Example

This example displays a message with the value of the first and second fields in the record that has just finished merging. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Word Application object.


```vb
Private Sub MailMergeApp_MailMergeAfterRecordMerge(ByVal Doc As Document) 
 
 With Doc.MailMerge.DataSource 
 MsgBox .DataFields(1).Value &; " " &; _ 
 .DataFields(2).Value &; " is finished merging." 
 End With 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

