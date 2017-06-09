---
title: TextRange2.Lines Property (PowerPoint)
ms.assetid: 09b52bda-e1ab-4cf2-bf38-25a156bf4c2e
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.Lines Property (PowerPoint)

Returns a TextRange2 object that represents the specified subset of text lines. Read-only.


## Syntax

 _expression_. **Lines**( **_Start_**, **_Length_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first line in the returned range.|
| _Length_|Optional|**Long**|The number of lines to be returned.|

### Return Value

TextRange2


## Remarks

If both Start and Length are omitted, the returned range starts with the first line and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one line.

If Length is specified but Start is omitted, the returned range starts with the first line in the specified range.

If Start is greater than the number of lines in the specified text, the returned range starts with the last line in the specified range.

If Length is greater than the number of lines from the specified starting line to the end of the text, the returned range contains all those lines.


## Example

This example formats as italic the first two lines of the second paragraph in shape two on slide one in the active PowerPoint presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _ 
 .TextFrame.TextRange2.Paragraphs(2) _ 
 .Lines(1, 2).Font.Italic = True 

```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


