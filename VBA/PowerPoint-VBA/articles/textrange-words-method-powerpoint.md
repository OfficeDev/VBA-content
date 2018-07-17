---
title: TextRange.Words Method (PowerPoint)
keywords: vbapp10.chm569012
f1_keywords:
- vbapp10.chm569012
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Words
ms.assetid: b8cd8dca-bf10-1041-dd9e-adc04b2df42d
ms.date: 06/08/2017
---


# TextRange.Words Method (PowerPoint)

Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the specified subset of text words.


## Syntax

 _expression_. **Words**( **_Start_**, **_Length_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first word in the returned range.|
| _Length_|Optional|**Long**|The number of words to be returned.|

### Return Value

TextRange


## Remarks

For information about counting or looping through the words in a text range, see the  **[TextRange](textrange-object-powerpoint.md)** object.

If both Start and Length are omitted, the returned range starts with the first word and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one word.

If Length is specified but Start is omitted, the returned range starts with the first word in the specified range.

If Start is greater than the number of words in the specified text, the returned range starts with the last word in the specified range.

If Length is greater than the number of words from the specified starting word to the end of the text, the returned range contains all those words.


## Example

This example formats as bold the second, third, and fourth words in the first paragraph in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _
    .TextFrame.TextRange.Paragraphs(1).Words(2, 3).Font _
    .Bold = True
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

