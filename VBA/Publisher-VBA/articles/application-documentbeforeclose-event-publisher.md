---
title: Application.DocumentBeforeClose Event (Publisher)
keywords: vbapb10.chm268435464
f1_keywords:
- vbapb10.chm268435464
ms.prod: publisher
api_name:
- Publisher.Application.DocumentBeforeClose
ms.assetid: d3ca4397-4df3-dc77-b758-d47e0bf13fe5
ms.date: 06/08/2017
---


# Application.DocumentBeforeClose Event (Publisher)

Occurs immediately before any open document closes.


## Syntax

 _expression_. **DocumentBeforeClose**( **_Doc_**,  **_Cancel_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The document that is being closed.|
|Cancel|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the document does not close when the procedure is finished.|

## Remarks

To access the  **Application** object events, declare an **Application** object variable in the General Declarations section of a code module. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft Publisher **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

This example prompts the user for a yes or no response before closing a document. To see this example work, this code must be placed in a class module and an instance of the class must be correctly initialized, using an example similar to the  **SetPubApp** routine below.


```vb
Private WithEvents PubApp As Application 
 
Sub SetPubApp() 
 Set PubApp = Publisher.Application 
End Sub 
 
Private Sub PubApp_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean) 
 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really want to close " _ 
 &; "the document?", vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

