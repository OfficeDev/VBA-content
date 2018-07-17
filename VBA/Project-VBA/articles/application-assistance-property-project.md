---
title: Application.Assistance Property (Project)
ms.prod: project-server
ms.assetid: f53bf107-9fd1-78f9-f8db-0b8c2acc5f72
ms.date: 06/08/2017
---


# Application.Assistance Property (Project)

 Gets an **Office.IAssistance** object that represents the Project Help system. Read-only **IAssistance**.


## Syntax

 _expression_. **Assistance**

 _expression_ A variable that represents an **Application** object.


## Remarks

For more information, see the  **IAssistance** object in the Microsoft Office Visual Basic Reference.


## Example

The following example displays the top-level page of the  **Project Help** window.


```vb
Sub ShowHelp()
    Dim theHelpSystem As Office.IAssistance
    
    Set theHelpSystem = Application.Assistance
    
    theHelpSystem.ShowHelp
End Sub
```


## Property value

 **<unknown type>**


