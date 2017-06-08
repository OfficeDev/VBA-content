---
title: TextRange2.Sentences Property (PowerPoint)
ms.assetid: 99dd9dd8-7d2b-4189-897e-1728f6ded6b4
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.Sentences Property (PowerPoint)

Returns a  **TextRange2** object that represents the specified subset of text sentences. Read-only.


## Syntax

 _expression_. **Sentences**( **_Start_**, **_Length_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first sentence in the returned range.|
| _Length_|Optional|**Long**|The number of sentences to be returned.|

### Return Value

TextRange2


## Remarks

If both Start and Length are omitted, the returned range starts with the first sentence and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one sentence.

If Length is specified but Start is omitted, the returned range starts with the first sentence in the specified range.

If Start is greater than the number of sentences in the specified text, the returned range starts with the last sentence in the specified range.

If Length is greater than the number of sentences from the specified starting sentence to the end of the text, the returned range contains all those sentences.


## Example

This example formats as bold the second sentence in the second paragraph in shape two on slide one in the active PowerPoint presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _ 
 .TextFrame.TextRange2.Paragraphs(2).Sentences(2).Font _ 
 .Bold = True 
 
```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


