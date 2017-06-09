---
title: Disable PAGE UP and PAGE DOWN Keys in a Form
ms.prod: access
ms.assetid: 998e1d00-f9d3-fcca-4535-390b0fd0d482
ms.date: 06/08/2017
---


# Disable PAGE UP and PAGE DOWN Keys in a Form

By default, the PAGE UP and PAGE DOWN keys can be used to navigate between records in a form. The followng example illustrates how to use a form's  **[KeyDown](form-keydown-event-access.md)** event to disable the use of the PAGE UP and PAGE DOWN keys in the form.


```vb
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 
 
    ' The Keycode value represents the key that 
    ' triggered the event. 
    Select Case KeyCode 
    
        ' Check for the PAGE UP and PAGE DOWN keys. 
        Case 33, 34 
 
        ' Cancel the keystroke. 
        KeyCode = 0 
    End Select 
End Sub
```


 **Note**  You must set the form's  **[KeyPreview](form-keypreview-property-access.md)** property to **True** in order for this procedure to work.


