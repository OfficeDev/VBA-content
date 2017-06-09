---
title: Worksheet.MailEnvelope Property (Excel)
keywords: vbaxl10.chm175150
f1_keywords:
- vbaxl10.chm175150
ms.prod: excel
api_name:
- Excel.Worksheet.MailEnvelope
ms.assetid: 9490f86c-a82f-d1ab-7315-29b89c799301
ms.date: 06/08/2017
---


# Worksheet.MailEnvelope Property (Excel)

Rrepresents an e-mail header for a document.


## Syntax

 _expression_ . **MailEnvelope**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example sets the comments for the header of the active worksheet.


```vb
Sub HeaderComments() 
 
 ActiveSheet.MailEnvelope.Introduction = "To Whom It May Concern: " 
 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

