---
title: Presentation.TitleMaster Property (PowerPoint)
keywords: vbapp10.chm583004
f1_keywords:
- vbapp10.chm583004
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.TitleMaster
ms.assetid: d5a84b2a-fff0-dcb5-e744-466428a586b5
ms.date: 06/08/2017
---


# Presentation.TitleMaster Property (PowerPoint)

Returns a  **[Master](master-object-powerpoint.md)** object that represents the title master for the specified presentation.


## Syntax

 _expression_. **TitleMaster**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Master


## Remarks

If the presentation doesn't have a title master, an error occurs.

Use the  **AddTitleMaster** method to add a title master to a presentation.


## Example

If the active presentation has a title master, this example sets the footer text for the title master.


```vb
With Application.ActivePresentation

    If .HasTitleMaster Then

        .TitleMaster.HeadersFooters.Footer.Text = "Introduction"

    End If

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

