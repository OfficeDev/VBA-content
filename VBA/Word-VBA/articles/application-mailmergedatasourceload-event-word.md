---
title: Application.MailMergeDataSourceLoad Event (Word)
keywords: vbawd10.chm4000020
f1_keywords:
- vbawd10.chm4000020
ms.prod: word
api_name:
- Word.Application.MailMergeDataSourceLoad
ms.assetid: 56158dbd-45df-76ef-260d-117becd2e9ac
ms.date: 06/08/2017
---


# Application.MailMergeDataSourceLoad Event (Word)

Occurs when the data source is loaded for a mail merge.


## Syntax

 _expression_ . **Private Sub object_MailMergeDataSourceLoad**( **_ByVal Doc As Document_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|

## Example

This example displays a message with the data source file name when the data source starts loading. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Word Application object.


```vb
Private Sub MailMergeApp_MailMergeDataSourceLoad(ByVal Doc As Document) 
 Dim strDSName As String 
 Dim intDSLength As Integer 
 Dim intDSStart As Integer 
 
 'Extract from the Name property only the file name 
 intDSLength = Len(Doc.MailMerge.DataSource.Name) 
 intDSStart = InStrRev(Doc.MailMerge.DataSource.Name, "\") 
 intDSStart = intDSLength - intDSStart 
 strDSName = Right(Doc.MailMerge.DataSource.Name, intDSStart) 
 
 'Deliver a message to user when data source is loading 
 MsgBox "Your data source, " &; strDSName &; ", is now loading." 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

