---
title: Application.WindowDeactivate Event (PowerPoint)
keywords: vbapp10.chm621010
f1_keywords:
- vbapp10.chm621010
ms.prod: powerpoint
api_name:
- PowerPoint.Application.WindowDeactivate
ms.assetid: 89bf2c09-a1a8-ed7f-74d5-49f8f7c027a7
ms.date: 06/08/2017
---


# Application.WindowDeactivate Event (PowerPoint)

Occurs when the application window or any document window is deactivated.


## Syntax

 _expression_. **WindowDeactivate**( **_Pres_**, **_Wn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation displayed in the deactivated window.|
| _Wn_|Required|**DocumentWindow**|The deactivated document window.|

## Example

This example finds the file name (without its extension) for the presentation in the window that is being deactivated. It then appends the .htm extension to the file name and saves it as a Web page in the same folder as the presentation.


```vb
Private Sub App_WindowDeactivate (ByVal Pres As Presentation, ByVal Wn As DocumentWindow)

    FindNum = InStr(1, Wn.Presentation.FullName, ".")

    If FindNum = 0 Then

        HTMLName = Wn.Presentation.FullName &; ".htm"

    Else

        HTMLName = Mid(Wn.Presentation.FullName, 1, FindNum - 1) _
            &; ".htm"

    End If

    Wn.Presentation.SaveCopyAs HTMLName, ppSaveAsHTML

    MsgBox "Presentation being saved in HTML format as " _
        &; HTMLName &; " ."

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

