---
title: MailMerge.SuppressBlankLines Property (Publisher)
keywords: vbapb10.chm6225927
f1_keywords:
- vbapb10.chm6225927
ms.prod: publisher
api_name:
- Publisher.MailMerge.SuppressBlankLines
ms.assetid: 3b41e0c0-8588-e86a-77ed-90c4692c03dc
ms.date: 06/08/2017
---


# MailMerge.SuppressBlankLines Property (Publisher)

 **True** to suppress blank lines when mail merge fields in a mail merge main document are empty. Read/write **Boolean**.


## Syntax

 _expression_. **SuppressBlankLines**

 _expression_A variable that represents a  **MailMerge** object.


### Return Value

Boolean


## Example

This example suppresses blank lines in the active publication when mail merge fields are blank. This example assumes that a mail merge data source is attached to the active publication.


```vb
Sub SuppressBlankLines() 
 ActiveDocument.MailMerge.SuppressBlankLines = True 
End Sub
```


