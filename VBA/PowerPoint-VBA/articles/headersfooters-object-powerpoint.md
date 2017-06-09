---
title: HeadersFooters Object (PowerPoint)
keywords: vbapp10.chm542000
f1_keywords:
- vbapp10.chm542000
ms.prod: powerpoint
api_name:
- PowerPoint.HeadersFooters
ms.assetid: 5fb10c90-0611-e797-836b-3f18b273af04
ms.date: 06/08/2017
---


# HeadersFooters Object (PowerPoint)

Contains all the  **[HeaderFooter](headerfooter-object-powerpoint.md)** objects on the specified slide, notes page, handout, or master.


## Remarks

Each  **HeaderFooter** object represents a header, footer, date and time, or slide number.


 **Note**   **HeaderFooter** objects aren't available for **[Slide](slide-object-powerpoint.md)** objects that represent notes pages. The **HeaderFooter** object that represents a header is available only for a notes master or handout master.


## Example

Use the  **[HeadersFooters](slide-headersfooters-property-powerpoint.md)** property to return the **HeadersFooters** object. Use the **[DateAndTime](headersfooters-dateandtime-property-powerpoint.md)**, **[Footer](headersfooters-footer-property-powerpoint.md)**, **[Header](headersfooters-header-property-powerpoint.md)**, or **[SlideNumber](headersfooters-slidenumber-property-powerpoint.md)** property to return an individual **HeaderFooter** object. The following example sets the footer text for slide one in the active presentation.


```vb
ActivePresentation.Slides(1).HeadersFooters.Footer _
    .Text = "Volcano Coffee"
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

