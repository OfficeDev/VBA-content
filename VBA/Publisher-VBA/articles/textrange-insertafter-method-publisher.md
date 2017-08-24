---
title: TextRange.InsertAfter Method (Publisher)
keywords: vbapb10.chm5308448
f1_keywords:
- vbapb10.chm5308448
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertAfter
ms.assetid: f647be29-68c7-b221-adf1-fa233583e74e
ms.date: 06/08/2017
---


# TextRange.InsertAfter Method (Publisher)

Returns a  **[TextRange](textrange-object-publisher.md)** object that represents text appended to the end of a text range.


## Syntax

 _expression_. **InsertAfter**( **_NewText_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|NewText|Required| **String**|The text to be inserted.|

### Return Value

TextRange


## Example

This example adds the Microsoft Publisher build number to the end of the first shape on the first page of the active publication. This example assumes the specified shape is a text frame and not another type of shape.


```vb
Sub AppendText() 
 With ActiveDocument.Pages(1).Shapes(1) 
 .TextFrame.TextRange.InsertAfter _ 
 NewText:="Microsoft Publisher Build : " &; Build 
 End With 
End Sub
```


