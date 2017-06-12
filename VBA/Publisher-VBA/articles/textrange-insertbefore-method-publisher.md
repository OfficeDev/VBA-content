---
title: TextRange.InsertBefore Method (Publisher)
keywords: vbapb10.chm5308449
f1_keywords:
- vbapb10.chm5308449
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertBefore
ms.assetid: b0e4355b-b1bc-ae78-08ad-000d577fd7db
ms.date: 06/08/2017
---


# TextRange.InsertBefore Method (Publisher)

Appends a string to the beginning of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.


## Syntax

 _expression_. **InsertBefore**( **_NewText_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|NewText|Required| **String**|The text to be inserted. The default value is an empty string.|

### Return Value

TextRange


## Example

This example adds the Microsoft Publisher build number and a paragraph break to the beginning of the first shape on the first page of the active publication. This example assumes the specified shape is a text frame and not another type of shape.


```vb
Sub InsertTextBefore() 
 With ActiveDocument.Pages(1).Shapes(1) 
 .TextFrame.TextRange.InsertBefore _ 
 NewText:="Microsoft Publisher Build : " &; Build &; vbCrLf 
 End With 
End Sub
```


