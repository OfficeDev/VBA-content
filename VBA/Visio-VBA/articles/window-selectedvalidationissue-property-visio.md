---
title: Window.SelectedValidationIssue Property (Visio)
keywords: vis_sdr.chm11662490
f1_keywords:
- vis_sdr.chm11662490
ms.prod: visio
api_name:
- Visio.Window.SelectedValidationIssue
ms.assetid: 7955338a-2a54-2726-a17a-81acc6bcfce7
ms.date: 06/08/2017
---


# Window.SelectedValidationIssue Property (Visio)

Gets or sets the validation issue that is selected in the  **Issues** window. Read/write.


## Syntax

 _expression_ . **SelectedValidationIssue**

 _expression_ A variable that represents a **[Window](window-object-visio.md)** object.


### Return Value

[ValidationIssue](validationissue-object-visio.md)


## Remarks

Attempting to get or set the  **SelectedValidationIssue** property on a window other than the **Issues** window, or when the **Issues** window is closed, returns an error.

If multiple issues are selected in the  **Issues** window, Visio returns the issue that has the focus.

If no issue is selected, Visio returns  **Nothing** . By default, issues that you have specified to be ignored are not displayed. If you set the property to **Nothing** , Visio clears the selection in the **Issues** window.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **SelectedValidationIssue** property to get the validation issue that is currently selected in the **Issues** window. If no issue is selected, the code displays a message box prompting the user to select an issue.


```vb
Set vsoIssuesWindow = Application.ActiveWindow.Windows.ItemFromID(Visio.VisWinTypes.visWinIDValidationIssues)
    
' If the Issues window is visible, find the selected validation issue.
    If vsoIssuesWindow.Visible Then
       Set vsoValidationIssue = vsoIssuesWindow.SelectedValidationIssue
    End If
    
' Respond to the case when no validation issue is selected. 
    If vsoValidationIssue Is Nothing Then
        MsgBox "Please select an issue."
    End If
```


