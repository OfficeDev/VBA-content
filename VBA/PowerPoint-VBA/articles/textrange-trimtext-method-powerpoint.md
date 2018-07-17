---
title: TextRange.TrimText Method (PowerPoint)
keywords: vbapp10.chm569016
f1_keywords:
- vbapp10.chm569016
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.TrimText
ms.assetid: 8566ed9d-c73a-d699-bcb7-edcd9a375afe
ms.date: 06/08/2017
---


# TextRange.TrimText Method (PowerPoint)

Returns a  **TextRange** object that represents the specified text minus any trailing spaces.


## Syntax

 _expression_. **TrimText**

 _expression_ A variable that represents a **TextRange** object.


### Return Value

TextRange


## Example

This example inserts the string " Text to trim " at the beginning of the text in shape two on slide one in the active presentation and then displays message boxes showing the string before and after it is trimmed.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2) _
        .TextFrame.TextRange
    With .InsertBefore("   Text to trim   ")
        MsgBox "Untrimmed: " &; """" &; .Text &; """"
        MsgBox "Trimmed: " &; """" &; .TrimText.Text &; """"
    End With
End With
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

