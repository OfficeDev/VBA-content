---
title: HeadersFooters.DateAndTime Property (PowerPoint)
keywords: vbapp10.chm542003
f1_keywords:
- vbapp10.chm542003
ms.prod: powerpoint
api_name:
- PowerPoint.HeadersFooters.DateAndTime
ms.assetid: 15d8f1a4-c48f-7afd-d701-d5e7545aadd4
ms.date: 06/08/2017
---


# HeadersFooters.DateAndTime Property (PowerPoint)

Returns a  **[HeaderFooter](headerfooter-object-powerpoint.md)** object that represents the date and time item that appears in the lower-left corner of a slide or in the upper-right corner of a notes page, handout, or outline. Read-only.


## Syntax

 _expression_. **DateAndTime**

 _expression_ A variable that represents a **HeadersFooters** object.


### Return Value

HeaderFooter


## Example

This example sets the date and time format for the slide master in the active presentation. This setting will apply to all slides that are based on this master.


```vb
Set myPres = Application.ActivePresentation

With myPres.SlideMaster.HeadersFooters.DateAndTime

    .Format = ppDateTimeMdyy

    .UseFormat = True

End With
```


## See also


#### Concepts


[HeadersFooters Object](headersfooters-object-powerpoint.md)

