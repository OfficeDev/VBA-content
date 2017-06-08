---
title: Application.MailMergeBeforeRecordMerge Event (Word)
keywords: vbawd10.chm4000019
f1_keywords:
- vbawd10.chm4000019
ms.prod: word
api_name:
- Word.Application.MailMergeBeforeRecordMerge
ms.assetid: ce7b6c4f-b100-32eb-440c-c557f7dd7340
ms.date: 06/08/2017
---


# Application.MailMergeBeforeRecordMerge Event (Word)

Occurs as a merge is executed for the individual records in a merge.


## Syntax

 _expression_ . **Private Sub object_MailMergeBeforeRecordMerge**( **_ByVal Doc As Document_** , **_Cancel As Boolean_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|
| _Cancel_|Required| **Boolean**| **True** stops the mail merge process, for the current record only, before it starts.|

## Example

This example verifies that the length of the postal code, which in this example is field number six, is fewer than five digits and, if it is, cancels the merge only for that record. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Microsoft Word Application object.


```vb
Private Sub MailMergeApp_MailMergeBeforeRecordMerge(ByVal _ 
 Doc As Document, Cancel As Boolean) 
 
 Dim intZipLength As Integer 
 
 intZipLength = Len(ActiveDocument.MailMerge _ 
 .DataSource.DataFields(6).Value) 
 
 'Cancel merge of this record only if 
 'the ZIP Code is fewer than five digits 
 If intZipLength < 5 Then 
 Cancel = True 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

