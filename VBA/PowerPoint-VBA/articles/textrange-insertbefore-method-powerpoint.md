---
title: TextRange.InsertBefore Method (PowerPoint)
keywords: vbapp10.chm569019
f1_keywords:
- vbapp10.chm569019
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.InsertBefore
ms.assetid: fbadcecd-a31b-8c8d-3281-63d70286bcff
ms.date: 06/08/2017
---


# TextRange.InsertBefore Method (PowerPoint)

Appends a string to the beginning of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.


## Syntax

 _expression_. **InsertBefore**( **_NewText_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewText_|Optional|**String**|The text to be appended. The default value is an empty string.|

## Example

This example appends the string "Test version: " to the beginning of the title on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(1)

    .TextFrame.TextRange.InsertBefore "Test version: "

End With
```

This example appends the contents of the Clipboard to the beginning of the title on slide one in the active presentation.




```vb
Application.ActivePresentation.Slides(1).Shapes(1).TextFrame _
    .TextRange.InsertBefore.Paste
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

