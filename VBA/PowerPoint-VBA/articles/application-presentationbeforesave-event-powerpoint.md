---
title: Application.PresentationBeforeSave Event (PowerPoint)
keywords: vbapp10.chm621018
f1_keywords:
- vbapp10.chm621018
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationBeforeSave
ms.assetid: 40943fe2-796f-45db-db0d-44b66854e196
ms.date: 06/08/2017
---


# Application.PresentationBeforeSave Event (PowerPoint)

Occurs before a presentation is saved.


## Syntax

 _expression_. **PresentationBeforeSave**( **_Pres_**, **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation being saved.|
| _Cancel_|Required|**Boolean**|**True** to cancel the save process.|

## Remarks

This event is triggered as the  **Save As** dialog box appears.

To access the  **Application** events, declare an **Application** variable in the General Declarations section of your code. Then set the variable equal to the **Application** object for which you want to access events. For information about using events with the Microsoft PowerPoint **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).


## Example

This example checks if there are revisions in a presentation and, if there are, asks whether to save the presentation. If a user's response is no, the save process is canceled. This example assumes an  **Application** object called PPTApp has been declared by using the **WithEvents** keyword.


```vb
Private Sub PPTApp_PresentationBeforeSave(ByVal Pres As Presentation, _
        Cancel As Boolean)

    Dim intResponse As Integer

    Set Pres = ActivePresentation

    If Pres.HasRevisionInfo Then

        intResponse = MsgBox(Prompt:="The presentation contains revisions. " &; _
            "Do you want to accept the revisions before saving?", Buttons:=vbYesNo)

        If intResponse = vbYes Then

            Cancel = True

            MsgBox "Your presentation was not saved."

        End If

    End If

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)