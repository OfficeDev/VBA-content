---
title: TextRange2.Lines Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Lines
ms.assetid: 5e20f089-c345-e22a-c136-483d13f7f658
ms.date: 06/08/2017
---


# TextRange2.Lines Property (Office)

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


```
Application.ActivePresentation.Slides(1).Shapes(2) _ 
 .TextFrame.TextRange2.Paragraphs(2) _ 
 .Lines(1, 2).Font.Italic = True 

```


## See also


#### Concepts


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

