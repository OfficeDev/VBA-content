---
title: TextRange.InsertAfter Method (PowerPoint)
keywords: vbapp10.chm569018
f1_keywords:
- vbapp10.chm569018
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.InsertAfter
ms.assetid: 2af4e134-c205-fbe6-a006-3fc1ca8d6a50
ms.date: 06/08/2017
---


# TextRange.InsertAfter Method (PowerPoint)

Appends a string to the end of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.


## Syntax

 _expression_. **InsertAfter**( **_NewText_** )

 _expression_ A variable that represents an **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewText_|Optional|**String**|The text to be inserted. The default value is an empty string.|

## Example

This example appends the string ": Test version" to the end of the title on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(1)

    .TextFrame.TextRange.InsertAfter ": Test version"

End With
```

This example appends the contents of the Clipboard to the end of the title on slide one.




```vb
Application.ActivePresentation.Slides(1).Shapes(1).TextFrame _
    .TextRange.InsertAfter.Paste
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

