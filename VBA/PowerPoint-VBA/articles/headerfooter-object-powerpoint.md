---
title: HeaderFooter Object (PowerPoint)
keywords: vbapp10.chm582000
f1_keywords:
- vbapp10.chm582000
ms.prod: powerpoint
api_name:
- PowerPoint.HeaderFooter
ms.assetid: 8aeafb02-adec-17c1-3108-565c78a64ed1
ms.date: 06/08/2017
---


# HeaderFooter Object (PowerPoint)

Represents a header, footer, date and time, slide number, or page number on a slide or master. All the  **HeaderFooter** objects for a slide or master are contained in a **[HeadersFooters](headersfooters-object-powerpoint.md)** object.


## Remarks

Use one of the properties listed in the following table to return the  **HeaderFooter** object.



|**Use this property**|**To return**|
|:-----|:-----|
|**[DateAndTime](headersfooters-dateandtime-property-powerpoint.md)**|A  **HeaderFooter** object that represents the date and time on the slide.|
|**[Footer](headersfooters-footer-property-powerpoint.md)**|A  **HeaderFooter** object that represents the footer for the slide.|
|**[Header](headersfooters-header-property-powerpoint.md)**|A  **HeaderFooter** object that represents the header for the slide. This works only for notes pages and handouts, not for slides.|
|**[SlideNumber](headersfooters-slidenumber-property-powerpoint.md)**|A  **HeaderFooter** object that represent the slide number (on a slide) or page number (on a notes page or a handout).|

 **Note**   **HeaderFooter** objects aren't available for **Slide** objects that represent notes pages. The **HeaderFooter** object that represents a header is available only for a notes master or handout master.


## Example

You can set properties of  **HeaderFooter** objects for single slides. The following example sets the footer text for slide one in the active presentation.


```vb
ActivePresentation.Slides(1).HeadersFooters.Footer.Text = "Volcano Coffee"
```

You can also set properties of  **HeaderFooter** objects for the slide master, title master, notes master, or handout master to affect all slides, title slides, notes pages, or handouts and outlines at the same time. The following example sets the text for the footer in the slide master for the active presentation, sets the format for the date and time, and turns on the display of slide numbers. These settings will apply to all slides that are based on this master that display master graphics and that have not had their footer and date and time set individually.




```vb
Set mySlidesHF = ActivePresentation.SlideMaster.HeadersFooters

With mySlidesHF

    .Footer.Visible = True

    .Footer.Text = "Regional Sales"

    .SlideNumber.Visible = True

    .DateAndTime.Visible = True

    .DateAndTime.UseFormat = True

    .DateAndTime.Format = ppDateTimeMdyy

End With
```

To clear header and footer information that has been set for individual slides and make sure all slides display the header and information you define for the slide master, run the following code before running the previous example.




```vb
For Each s In ActivePresentation.Slides

    s.DisplayMasterShapes = True

    s.HeadersFooters.Clear

Next
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

