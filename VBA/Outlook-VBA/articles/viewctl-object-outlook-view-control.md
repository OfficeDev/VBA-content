---
title: ViewCtl Object (Outlook View Control)
ms.prod: outlook
ms.assetid: e5737688-6196-bc0a-767c-7b1fe7071fce
ms.date: 06/08/2017
---


# ViewCtl Object (Outlook View Control)

Displays information about a specific folder and can be integrated into a Microsoft Outlook form or folder home page that provides access to Outlook data.


## Remarks

The  **ViewCtl** object provides programmatic access to the View Control. Use this control only within Outlook, in an HTML folder home page that is hosted in Outlook, or a custom Outlook form that is displayed by an Outlook add-in. This makes sure that Outlook is running and that the View Control is not subject to other factors that may adversely affect the Outlook process from continuing to be available for the View Control's use. Do not use the View Control in any scenario outside the Outlook process, such as in an HTML page hosted in a browser. Out-of-process scenarios are not supported. For more information, see [Known issues when you use Outlook View Control with Outlook 2010](http://support.microsoft.com/kb/2511230).

You can set the control's properties programmatically to customize the folder and the view that is displayed in the control. You can use the control to create a variety of solutions that integrate Outlook data.

For example, you can place multiple View Controls on an HTML page so that a user can view the contents of more than one folder in a single window. This can be useful in a scenario where you want to display calendar information for more than one user at a time.

To use the  **ViewCtl** object in your code, you must set a reference to the View Control's type library.

To set a reference to the View Control's type library:


1. In the Visual Basic for Applications Code Editor, on the  **Tools** menu, click **References**. The  **References** command on the **Tools** menu is available only when a **Module** window is open and active in **Design** view.
    
2. Select the  **Microsoft Outlook View Control** check box.
    

