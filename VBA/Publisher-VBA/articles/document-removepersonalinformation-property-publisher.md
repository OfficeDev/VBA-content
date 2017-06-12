---
title: Document.RemovePersonalInformation Property (Publisher)
keywords: vbapb10.chm196742
f1_keywords:
- vbapb10.chm196742
ms.prod: publisher
api_name:
- Publisher.Document.RemovePersonalInformation
ms.assetid: bbc1aee1-90ca-966e-c17c-579064318cd1
ms.date: 06/08/2017
---


# Document.RemovePersonalInformation Property (Publisher)

Returns or sets a  **Boolean** that represents whether to save personal information when the file is saved. Read/write.


## Syntax

 _expression_. **RemovePersonalInformation**

 _expression_A variable that represents a  **Document** object.


### Return Value

Boolean


## Remarks

The information removed from the document is the Author, Manager, Company, and the GUID of the computer on which the document was created.

The default setting for this property is  **False**.


## Example

This example removes the personal information from the active document.


```vb
ActiveDocument.RemovePersonalInformation = True 

```


