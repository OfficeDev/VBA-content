---
title: HeadersFooters.Footer Property (PowerPoint)
keywords: vbapp10.chm542006
f1_keywords:
- vbapp10.chm542006
ms.prod: powerpoint
api_name:
- PowerPoint.HeadersFooters.Footer
ms.assetid: a684ae25-7174-50e5-095e-0073873628e6
ms.date: 06/08/2017
---


# HeadersFooters.Footer Property (PowerPoint)

Returns a  **[HeaderFooter](headerfooter-object-powerpoint.md)** object that represents the footer that appears at the bottom of a slide or in the lower-left corner of a notes page, handout, or outline. Read-only.


## Syntax

 _expression_. **Footer**

 _expression_ A variable that represents a **HeadersFooters** object.


### Return Value

HeaderFooter


## Example

This example sets the text for the footer on the slide master in the active presentation and sets the footer, date and time, and slide number to appear on the title slide.


```vb
With Application.ActivePresentation.SlideMaster.HeadersFooters

    .Footer.Text = "Introduction"

    .DisplayOnTitleSlide = True

End With
```


## See also


#### Concepts


[HeadersFooters Object](headersfooters-object-powerpoint.md)

