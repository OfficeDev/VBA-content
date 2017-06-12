---
title: TextRange.Paragraphs Method (PowerPoint)
keywords: vbapp10.chm569010
f1_keywords:
- vbapp10.chm569010
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Paragraphs
ms.assetid: 5062eccf-4db2-692f-501e-b7d214181171
ms.date: 06/08/2017
---


# TextRange.Paragraphs Method (PowerPoint)

Returns a  **TextRange** object that represents the specified subset of text paragraphs.


## Syntax

 _expression_. **Paragraphs**( **_Start_**, **_Length_** )

 _expression_ An expression that returns a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first paragraph in the returned range.|
| _Length_|Optional|**Long**|The number of paragraphs to be returned.|

### Return Value

TextRange


## Remarks

For information about counting or looping through the paragraphs in a text range, see the  **[TextRange](textrange-object-powerpoint.md)** object.

If both Start and Length are omitted, the returned range starts with the first paragraph and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one paragraph.

If Length is specified but Start is omitted, the returned range starts with the first paragraph in the specified range.

If Start is greater than the number of paragraphs in the specified text, the returned range starts with the last paragraph in the specified range.

If Length is greater than the number of paragraphs from the specified starting paragraph to the end of the text, the returned range contains all those paragraphs.


## Example

This example formats as italic the first two lines of the second paragraph in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _
    .TextFrame.TextRange.Paragraphs(2) _
    .Lines(1, 2).Font.Italic = True
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

