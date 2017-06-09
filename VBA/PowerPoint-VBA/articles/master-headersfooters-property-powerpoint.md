---
title: Master.HeadersFooters Property (PowerPoint)
keywords: vbapp10.chm533004
f1_keywords:
- vbapp10.chm533004
ms.prod: powerpoint
api_name:
- PowerPoint.Master.HeadersFooters
ms.assetid: ac9f3282-32be-c561-e5cb-80e35db1797d
ms.date: 06/08/2017
---


# Master.HeadersFooters Property (PowerPoint)

Returns a  **[HeadersFooters](headersfooters-object-powerpoint.md)** collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. Read-only.


## Syntax

 _expression_. **HeadersFooters**

 _expression_ A variable that represents a **Master** object.


### Return Value

HeadersFooters


## Example

This example sets the footer text and the date and time format for the notes master in the active presentation and sets the date and time to be updated automatically.


```vb
With ActivePresentation.NotesMaster.HeadersFooters

    .Footer.Text = "Regional Sales"

    With .DateAndTime

        .UseFormat = True

        .Format = ppDateTimeHmmss

    End With

End With
```


## See also


#### Concepts


[Master Object](master-object-powerpoint.md)

