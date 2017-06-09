---
title: "How to: Resize a Control Horizontally and Vertically with the Form"
keywords: olfm10.chm3077352
f1_keywords:
- olfm10.chm3077352
ms.prod: outlook
ms.assetid: 67dfbd5d-98a8-41d1-a92f-56ae381d2e50
ms.date: 06/08/2017
---


# How to: Resize a Control Horizontally and Vertically with the Form

The following code sample shows how to resize a control with a form. It uses the  **[OlkControl](olkcontrol-object-outlook.md)** class that represents some basic properties (for example, **[HorizontalLayout](olkcontrol-horizontallayout-property-outlook.md)** and ** [VerticalLayout](olkcontrol-verticallayout-property-outlook.md)**) common to Outlook form controls. It assumes an existing Outlook text box control, myTextBox, in the form, and uses casting in Visual Basic to allow the text box control to use the properties of  **OlkControl**.


```vb
Dim olkCtrl As Outlook.OlkControl

    ' Let the text box control use the properties of OlkControl
    Set olkCtrl = myTextBox
    
    ' Enable automatic adjustments of the layout with respect to the rest of the form
    olkCtrl.EnableAutoLayout = True

    ' Allow resizing the text box control horizontally and vertically with the form
    olkCtrl.HorizontalLayout = olHorizontalLayoutGrow
    olkCtrl.VerticalLayout = olVerticalLayoutGrow

```


