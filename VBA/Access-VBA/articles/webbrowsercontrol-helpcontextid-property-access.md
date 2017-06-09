---
title: WebBrowserControl.HelpContextId Property (Access)
keywords: vbaac10.chm14684
f1_keywords:
- vbaac10.chm14684
ms.prod: access
api_name:
- Access.WebBrowserControl.HelpContextId
ms.assetid: f0678d0c-eb24-2398-208f-971772ea2c21
ms.date: 06/08/2017
---


# WebBrowserControl.HelpContextId Property (Access)

The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.


## Syntax

 _expression_. **HelpContextId**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks


 **Note**  If you enter the context ID of the Help file topic as a positive number, the help topic will display in a "full" help topic window. If you add a minus sign ("-") in front of the context ID, the help topic will be displayed in a "pop-up" window. It is important to note the context ID does not have to have a negative number when authored in Microsoft Help Workshop. You must add the minus sign when setting the property to make the topic display in the pop-up window.

You can create a custom Help file to document forms, reports, or applications you create with Microsoft Access.

If a control's  **HelpContextID** property setting is 0 (the default), Microsoft Access uses the form's **HelpContextID** and **HelpFile** properties to identify the Help topic to display. If you press F1 in a view other than Form view or if the **HelpContextID** property setting for both the form and the control is 0, a Microsoft Access Help topic is displayed.


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

