---
title: NavigationButton.HelpContextId Property (Access)
keywords: vbaac10.chm10481
f1_keywords:
- vbaac10.chm10481
ms.prod: access
api_name:
- Access.NavigationButton.HelpContextId
ms.assetid: 6549fc57-0f33-fe6b-ee5e-88b39107c035
ms.date: 06/08/2017
---


# NavigationButton.HelpContextId Property (Access)

The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.


## Syntax

 _expression_. **HelpContextId**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks


 **Note**  If you enter the context ID of the Help file topic as a positive number, the help topic will display in a "full" help topic window. If you add a minus sign ("-") in front of the context ID, the help topic will be displayed in a "pop-up" window. It is important to note the context ID does not have to have a negative number when authored in Microsoft Help Workshop. You must add the minus sign when setting the property to make the topic display in the pop-up window.

You can create a custom Help file to document forms, reports, or applications you create with Microsoft Access.

When you press the F1 key in Form view, Microsoft Access calls the Microsoft Help Workshop or Microsoft HTML Help Workshop application, loads the custom Help file specified by the  **HelpFile** property setting for the form or report, and displays the Help topic specified by the **HelpContextID** property setting.

If a control's  **HelpContextID** property setting is 0 (the default), Microsoft Access uses the form's **HelpContextID** and **HelpFile** properties to identify the Help topic to display. If you press F1 in a view other than Form view or if the **HelpContextID** property setting for both the form and the control is 0, a Microsoft Access Help topic is displayed.


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)

