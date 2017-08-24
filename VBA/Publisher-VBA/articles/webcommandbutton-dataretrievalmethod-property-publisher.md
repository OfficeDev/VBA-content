---
title: WebCommandButton.DataRetrievalMethod Property (Publisher)
keywords: vbapb10.chm3932166
f1_keywords:
- vbapb10.chm3932166
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.DataRetrievalMethod
ms.assetid: 81b89a3b-dcc5-c2b5-fbc4-6e02b587bc42
ms.date: 06/08/2017
---


# WebCommandButton.DataRetrievalMethod Property (Publisher)

Sets or returns a  **PbSubmitDataRetrievalMethodType** that represents the way data from a Web form is processed. Read/write.


## Syntax

 _expression_. **DataRetrievalMethod**

 _expression_A variable that represents a  **WebCommandButton** object.


### Return Value

PbSubmitDataRetrievalMethodType


## Remarks

The  **DataRetrievalMethod** property value can be one of the **[PbSubmitDataRetrievalMethodType](pbsubmitdataretrievalmethodtype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example sets Microsoft Publisher to process data on the Web form in the current publication by sending an e-mail message to a specified e-mail address.


```vb
Sub WebFormData() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalEmail 
 .EmailAddress = "someone@example.com" 
 .EmailSubject = "Web form data" 
 End With 
End Sub
```


