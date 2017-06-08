---
title: Presentation.NotesMaster Property (PowerPoint)
keywords: vbapp10.chm583009
f1_keywords:
- vbapp10.chm583009
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.NotesMaster
ms.assetid: 0889b69b-4c51-82cf-ccc2-ccb211d8a34e
ms.date: 06/08/2017
---


# Presentation.NotesMaster Property (PowerPoint)

Returns a  **[Master](master-object-powerpoint.md)** object that represents the notes master. Read-only.


## Syntax

 _expression_. **NotesMaster**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Master


## Example

This example sets the header and footer text for the notes master for the active presentation.


```vb
With Application.ActivePresentation. NotesMaster.HeadersFooters 
    .Header.Text = "Employee Guidelines" 
    .Footer.Text = "Volcano Coffee" 
End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

