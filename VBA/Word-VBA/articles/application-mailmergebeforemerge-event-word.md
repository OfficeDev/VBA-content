---
title: Application.MailMergeBeforeMerge Event (Word)
keywords: vbawd10.chm4000018
f1_keywords:
- vbawd10.chm4000018
ms.prod: word
api_name:
- Word.Application.MailMergeBeforeMerge
ms.assetid: 968cf799-255f-b6fc-f576-7aec093ab1cb
ms.date: 06/08/2017
---


# Application.MailMergeBeforeMerge Event (Word)

Occurs when a merge is executed before any records merge.


## Syntax

 _expression_ . **Private Sub object_MailMergeBeforeMerge**( **_ByVal Doc As Document_** , **_ByVal StartRecord As Long_** , **_ByVal EndRecord As Long_** , **_Cancel As Boolean_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|
| _StartRecord_|Required| **Long**|The first record in the data source to include in the mail merge.|
| _EndRecord_|Required| **Long**|The last record in the data source to include in the mail merge.|
| _EndRecord_|Required| **Boolean**| **True** stops the mail merge process before it starts.|

## Example

This example displays a message before the mail merge process begins, asking the user if they want to continue. If the user clicks No, the merge process is canceled. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Microsoft Word Application object.


```vb
Private Sub MailMergeApp_MailMergeBeforeMerge(ByVal Doc As Document, _ 
 ByVal StartRecord As Long, ByVal EndRecord As Long, _ 
 Cancel As Boolean) 
 
 Dim intVBAnswer As Integer 
 
 'Request whether the user wants to continue with the merge 
 intVBAnswer = MsgBox("Mail Merge for " &; _ 
 Doc.Name &; " is now starting. " &; _ 
 "Do you want to continue?", vbYesNo, "MailMergeBeforeMerge Event") 
 
 'If users response to question is No, cancel the merge process 
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


[Application Object](application-object-word.md)

