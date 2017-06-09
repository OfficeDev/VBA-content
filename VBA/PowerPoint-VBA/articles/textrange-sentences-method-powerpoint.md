---
title: TextRange.Sentences Method (PowerPoint)
keywords: vbapp10.chm569011
f1_keywords:
- vbapp10.chm569011
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Sentences
ms.assetid: c3640cb8-f78a-2934-bbe0-506cb8d2534c
ms.date: 06/08/2017
---


# TextRange.Sentences Method (PowerPoint)

Returns a  **TextRange** object that represents the specified subset of text sentences.


## Syntax

 _expression_. **Sentences**( **_Start_**, **_Length_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first sentence in the returned range.|
| _Length_|Optional|**Long**|The number of sentences to be returned.|

### Return Value

TextRange


## Remarks

If both Start and Length are omitted, the returned range starts with the first sentence and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one sentence.

If Length is specified but Start is omitted, the returned range starts with the first sentence in the specified range.

If Start is greater than the number of sentences in the specified text, the returned range starts with the last sentence in the specified range.

If Length is greater than the number of sentences from the specified starting sentence to the end of the text, the returned range contains all those sentences.

For information about counting or looping through the sentences in a text range, see the  **[TextRange](textrange-object-powerpoint.md)** object.


## Example

This example formats as bold the second sentence in the second paragraph in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _
    .TextFrame.TextRange.Paragraphs(2).Sentences(2).Font _
    .Bold = True
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

