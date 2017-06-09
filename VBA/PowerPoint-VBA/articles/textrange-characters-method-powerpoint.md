---
title: TextRange.Characters Method (PowerPoint)
keywords: vbapp10.chm569013
f1_keywords:
- vbapp10.chm569013
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Characters
ms.assetid: 019c15d3-349d-ab10-7448-70bf81176150
ms.date: 06/08/2017
---


# TextRange.Characters Method (PowerPoint)

Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the specified subset of text characters. For information about counting or looping through the characters in a text range, see the **[TextRange](textrange-object-powerpoint.md)** object.


## Syntax

 _expression_. **Characters**( **_Start_**, **_Length_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first character in the returned range.|
| _Length_|Optional|**Long**|The number of characters to be returned.|

### Return Value

TextRange


## Remarks

If both Start and Length are omitted, the returned range starts with the first character and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one character.

If Length is specified but Start is omitted, the returned range starts with the first character in the specified range.

If Start is greater than the number of characters in the specified text, the returned range starts with the last character in the specified range.

If Length is greater than the number of characters from the specified starting character to the end of the text, the returned range contains all those characters.


## Example

This example sets the text for shape two on slide one in the active presentation and then makes the second character a subscript character with a 20-percent offset.


```vb
Dim charRange As TextRange

With Application.ActivePresentation.Slides(1).Shapes(2)

    Set charRange = .TextFrame.TextRange.InsertBefore("H2O")

    charRange.Characters(2).Font.BaselineOffset = -0.2

End With


```

This example formats every subscript character in shape two on slide one as bold.




```vb
With Application.ActivePresentation.Slides(1).Shapes(2) _
    .TextFrame.TextRange

    For i = 1 To .Characters.Count

        With .Characters(i).Font

            If .Subscript Then .Bold = True

        End With

    Next

End With


```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

