---
title: Application.NewDocument Event (Publisher)
keywords: vbapb10.chm268435462
f1_keywords:
- vbapb10.chm268435462
ms.prod: publisher
api_name:
- Publisher.Application.NewDocument
ms.assetid: 629cf55c-5134-4207-14df-143b517b9f36
ms.date: 06/08/2017
---


# Application.NewDocument Event (Publisher)

Occurs when a new publication is created.


## Syntax

 _expression_. **NewDocument**( **_Doc_**, )

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The new document.|

## Example

This example displays a message when a new publication is created.


```vb
Private Sub appPub_NewDocument(ByVal Doc As Document) 
 MsgBox "This is a new publication." 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

