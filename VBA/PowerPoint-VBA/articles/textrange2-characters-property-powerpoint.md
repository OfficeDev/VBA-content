---
title: TextRange2.Characters Property (PowerPoint)
ms.assetid: 5a7f47ac-8d22-42f3-af2f-0a381315ce07
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.Characters Property (PowerPoint)

Read-only.


## Syntax

 _expression_. **Characters**( **_Start_**, **_Length_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first character in the returned range.|
| _Length_|Optional|**Long**|The number of characters to be returned.|

### Return Value

TextRange2


## Remarks

If both Start and Length are omitted, the returned range starts with the first character and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one character.

If Length is specified but Start is omitted, the returned range starts with the first character in the specified range.

If Start is greater than the number of characters in the specified text, the returned range starts with the last character in the specified range.

If Length is greater than the number of characters from the specified starting character to the end of the text, the returned range contains all those characters.


## Example

This example sets the text for shape two on slide one in the active presentation and then makes the second character a subscript character with a 20-percent offset.


```vb
Dim charRange As TextRange2 
With Application.ActivePresentation.Slides(1).Shapes(2) 
 Set charRange = .TextFrame.TextRange2.InsertBefore("H2O") 
 charRange.Characters(2).Font.BaselineOffset = -0.2 
End With 

```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


