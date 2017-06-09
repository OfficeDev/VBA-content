---
title: SlideRange.HeadersFooters Property (PowerPoint)
keywords: vbapp10.chm532004
f1_keywords:
- vbapp10.chm532004
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.HeadersFooters
ms.assetid: 204e867b-af78-81ad-bcc3-aa0e77d36a36
ms.date: 06/08/2017
---


# SlideRange.HeadersFooters Property (PowerPoint)

Returns a  **[HeadersFooters](headersfooters-object-powerpoint.md)** collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. Read-only.


## Syntax

 _expression_. **HeadersFooters**

 _expression_ A variable that represents a **SlideRange** object.


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


[SlideRange Object](sliderange-object-powerpoint.md)

