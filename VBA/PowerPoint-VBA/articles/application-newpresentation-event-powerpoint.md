---
title: Application.NewPresentation Event (PowerPoint)
keywords: vbapp10.chm621007
f1_keywords:
- vbapp10.chm621007
ms.prod: powerpoint
api_name:
- PowerPoint.Application.NewPresentation
ms.assetid: 63a6a83d-74c4-88ac-4972-d54907f5af8a
ms.date: 06/08/2017
---


# Application.NewPresentation Event (PowerPoint)

Occurs after a presentation is created, as it is added to the  **[Presentations](presentations-object-powerpoint.md)** collection.


## Syntax

 _expression_. **NewPresentation**( **_Pres_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The new presentation.|

## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this event maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.EApplication_NewPresentationEventHandler** (the **NewPresentation** delegate.)
    
-  **Microsoft.Office.Interop.PowerPoint.EApplication_Event.NewPresentation** (the **NewPresentation** event.)
    

## Example

This example uses the  **RGB** function to set the slide master background color for the new presentation to salmon pink and then applies the third color scheme to the new presentation.


```vb
Private Sub App_NewPresentation(ByVal Pres As Presentation) 
    With Pres 
        Set CS3 = .ColorSchemes(3) 
        CS3.Colors(ppBackground).RGB = RGB(240, 115, 100) 
        .SlideMaster.ColorScheme = CS3 
    End With 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

