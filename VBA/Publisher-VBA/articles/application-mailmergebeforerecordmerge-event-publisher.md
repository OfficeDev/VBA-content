---
title: Application.MailMergeBeforeRecordMerge Event (Publisher)
keywords: vbapb10.chm268435474
f1_keywords:
- vbapb10.chm268435474
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeBeforeRecordMerge
ms.assetid: 67ae8255-336d-0ff8-7927-fbd31262c115
ms.date: 06/08/2017
---


# Application.MailMergeBeforeRecordMerge Event (Publisher)

Occurs as a merge is executed for the individual records in a merge.


## Syntax

 _expression_. **MailMergeBeforeRecordMerge**( **_Doc_**,  **_Cancel_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The mail merge main document.|
|Cancel|Required| **Boolean**|Stops the mail merge process, for the current record only, before it starts.|

## Remarks

To access the  **Application** object events, declare an **Application** object variable in the General Declarations section of a code module. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft Publisher **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

This example verifies that the length of the ZIP Code (which in this example is field number six) is less than five and if it is, cancels the merge for that record only.


```vb
Private Sub MailMergeApp_MailMergeBeforeRecordMerge(ByVal _ 
 Doc As Document, Cancel As Boolean) 
 
 Dim intZipLength As Integer 
 
 intZipLength = Len(ActiveDocument.MailMerge _ 
 .DataSource.DataFields(6).Value) 
 
 'Cancel merge of this record only if 
 'the ZIP Code has fewer than five digits 
 If intZipLength < 5 Then 
 Cancel = True 
 End If 
 
End Sub
```

For this event to occur, you must place the following line of code in the global declarations section of your module and run the following initialization routine.




```vb
Private WithEvents MailMergeApp As Application 
 
Sub InitializeMailMergeApp() 
 Set MailMergeApp = Publisher.Application 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

