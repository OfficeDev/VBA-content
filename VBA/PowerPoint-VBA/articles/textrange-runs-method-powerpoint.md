---
title: TextRange.Runs Method (PowerPoint)
keywords: vbapp10.chm569015
f1_keywords:
- vbapp10.chm569015
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Runs
ms.assetid: 0bf2724a-0735-bd79-31e5-894d1320b9b2
ms.date: 06/08/2017
---


# TextRange.Runs Method (PowerPoint)

Returns a  **TextRange** object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes.


## Syntax

 _expression_. **Runs**( **_Start_**, **_Length_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first run in the returned range.|
| _Length_|Optional|**Long**|The number of runs to be returned.|

### Return Value

TextRange


## Remarks

If both Start and Length are omitted, the returned range starts with the first run and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one run.

If Length is specified but Start is omitted, the returned range starts with the first run in the specified range.

If Start is greater than the number of runs in the specified text, the returned range starts with the last run in the specified range.

If Length is greater than the number of runs from the specified starting run to the end of the text, the returned range contains all those runs.

A run consists of all characters from the first character after a font change to the second-to-last character that has the same font attributes. For example, consider the following sentence:

This italic word is not  **bold**.

In the preceding sentence, the first run consists of the word "This" only if the space after the word "This" isn't formatted as italic (if the space is italic, the first run is only the first three characters, or "Thi"). Likewise, the second run contains the word "italic" only if the space after the word is formatted as italic.

For information about counting or looping through the runs in a text range, see the  **[TextRange](textrange-object-powerpoint.md)** object.


## Example

This example formats the second run in shape two on slide one in the active presentation as bold italic if it is already italic.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2) _
        .TextFrame.TextRange

    With .Runs(2).Font
        If .Italic Then
            .Bold = True
        End If
    End With

End With


```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

