---
title: Application.MailMergeAfterRecordMerge Event (Publisher)
keywords: vbapb10.chm268435472
f1_keywords:
- vbapb10.chm268435472
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeAfterRecordMerge
ms.assetid: 550c3310-01ba-718f-4c1d-cbf3ce077d27
ms.date: 06/08/2017
---


# Application.MailMergeAfterRecordMerge Event (Publisher)

Occurs after each record in the data source successfully merges in a mail merge.


## Syntax

 _expression_. **MailMergeAfterRecordMerge**( **_Doc_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The mail merge main document.|

## Remarks

If you maintain a customer management database, you can use the  **MailMergeAfterRecordMerge** event to update the database for each merged record.

To access the  **Application** object events, declare an **Application** object variable in the General Declarations section of a code module. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft Publisher **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

This example displays a message with the value of the first and second fields in the record that has just finished merging.


```vb
Private Sub MailMergeApp_MailMergeAfterRecordMerge(ByVal Doc As Document) 
 
 With ActiveDocument.MailMerge.DataSource 
 MsgBox .DataFields.Item(3).Value &; " " &; _ 
 .DataFields.Item(2).Value &; " is finished merging." 
 End With 
 
End Sub
```

For this event to occur, you must place the following line of code in the General Declarations section of your module and run the following initialization routine.




```vb
Private WithEvents MailMergeApp As Application 
 
Sub InitializeMailMergeApp() 
 Set MailMergeApp = Publisher.Application 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

