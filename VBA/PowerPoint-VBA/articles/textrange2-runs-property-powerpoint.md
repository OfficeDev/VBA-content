---
title: TextRange2.Runs Property (PowerPoint)
ms.assetid: 1799ac12-3ebb-4790-a433-9b1f27ecdb38
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.Runs Property (PowerPoint)

Gets a  **TextRange2** object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes. Read-only.


## Syntax

 _expression_. **Runs**( **_Start_**, **_Length_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Long**|The first run in the returned range.|
| _Length_|Optional|**Long**|The number of runs to be returned.|

### Return Value

TextRange2


## Remarks

If both  _Start_ and _Length_ are omitted, the returned range starts with the first run and ends with the last paragraph in the specified range.

If  _Start_ is specified but _Length_ is omitted, the returned range contains one run.

If  _Length_ is specified but _Start_ is omitted, the returned range starts with the first run in the specified range.

If  _Start_ is greater than the number of runs in the specified text, the returned range starts with the last run in the specified range.

If  _Length_ is greater than the number of runs from the specified starting run to the end of the text, the returned range contains all those runs.

A run consists of all characters from the first character after a font change to the second-to-last character with the same font attributes. For example, consider the following sentence:

This  _italic_ word is not bold.

In the preceding sentence, the first run consists of the word "This" only if the space after the word "This" isn't formatted as italic (if the space is italic, the first run is only the first three characters, or "Thi"). Likewise, the second run contains the word "italic" only if the space after the word is formatted as italic.


## Example

This example formats the second run in shape two on slide one in the active presentation as bold italic if it's already italic.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2) _ 
        .TextFrame.TextRange2 
    With .Runs(2).Font 
        If .Italic Then 
            .Bold = True 
        End If 
    End With 
End With

```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


