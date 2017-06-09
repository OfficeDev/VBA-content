---
title: Application.ColorSchemeChanged Event (PowerPoint)
keywords: vbapp10.chm621017
f1_keywords:
- vbapp10.chm621017
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ColorSchemeChanged
ms.assetid: 8b517ce7-879d-bb96-477b-072477c991d5
ms.date: 06/08/2017
---


# Application.ColorSchemeChanged Event (PowerPoint)

Occurs after a color scheme is changed.


## Syntax

 _expression_. **ColorSchemeChanged**( **_SldRange_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SldRange_|Required|**SlideRange**| The range of slides affected by the change.|

## Remarks

Actions which trigger this event would include actions such as modifying a slide's or slide master's color scheme, or applying a template.

To access the  **Application** events, declare an **Application** variable in the General Declarations section of your code. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft PowerPoint **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).


## Example

This example displays a message when the color scheme for the selected slide or slides is changed. This example assumes an  **Application** object called PPTApp has been declared by using the **WithEvents** keyword.


```vb
Private Sub PPTApp_ColorSchemeChanged(ByVal SldRange As SlideRange)



    If SldRange.Count = 1 Then

        MsgBox "You've changed the color scheme for " _
            &; SldRange.Name &; "."

    Else

        MsgBox "You've changed the color scheme for " _
            &; SldRange.Count &; " slides."

    End If

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

