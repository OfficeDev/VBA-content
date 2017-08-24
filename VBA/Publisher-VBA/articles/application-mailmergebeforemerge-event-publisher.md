---
title: Application.MailMergeBeforeMerge Event (Publisher)
keywords: vbapb10.chm268435473
f1_keywords:
- vbapb10.chm268435473
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeBeforeMerge
ms.assetid: 735ef282-e99f-b3f2-c509-b180bea30d36
ms.date: 06/08/2017
---


# Application.MailMergeBeforeMerge Event (Publisher)

Occurs when a merge is executed before any records in a mail merge have merged.


## Syntax

 _expression_. **MailMergeBeforeMerge**( **_Doc_**,  **_StartRecord_**,  **_EndRecord_**,  **_Cancel_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The mail merge main document.|
|StartRecord|Required| **Long**|The first record in the data source to include in the mail merge.|
|EndRecord|Required| **Long**|The last record in the data source to include in the mail merge.|
|Cancel|Required| **Boolean**|Stops the mail merge process before it starts.|

## Remarks

To access the  **Application** object events, declare an **Application** object variable in the General Declarations section of a code module. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft Publisher **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

This example displays a message before the mail merge process begins, asking the user if they want to continue. If the user clicks No, the merge process is canceled.


```vb
Private Sub MailMergeApp_MailMergeBeforeMerge(ByVal Doc As Document, _ 
 ByVal StartRecord As Long, ByVal EndRecord As Long, _ 
 Cancel As Boolean) 
 
 Dim intVBAnswer As Integer 
 
 Set Doc = ActiveDocument 
 
 'Request whether the user wants to continue with the merge 
 intVBAnswer = MsgBox("Mail Merge for " &; Doc.Name &; _ 
 " is now starting. Do you want to continue?", _ 
 vbYesNo, "Event!") 
 
 'If user's response to question is No, then cancel merge process 
 'and deliver a message to the user stating the merge is canceled 
 If intVBAnswer = vbNo Then 
 Cancel = True 
 MsgBox "You have canceled mail merge for " &; _ 
 Doc.Name &; "." 
 End If 
 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

