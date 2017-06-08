---
title: Detecting Errors
ms.prod: outlook
ms.assetid: 73778714-906c-a57a-00d8-6450bfc9a6d9
ms.date: 06/08/2017
---


# Detecting Errors

The custom user interface XML markup that you return in the  [IRibbonExtensibility.GetCustomUI](http://msdn.microsoft.com/library/a0106415-999e-94da-379c-70fb7aa6119f%28Office.15%29.aspx) call typically contains callbacks that run when the corresponding Microsoft Office Fluent user interface (UI) that you are customizing (for example, explorer, inspector, or menu) is about to be displayed. 

You must add each callback in your XML markup to the add-in class that implements [IRibbonExtensibility](http://msdn.microsoft.com/library/b27a7576-b6f5-031e-e307-78ef5f8507e0%28Office.15%29.aspx). In addition, you must declare the callbacks as public procedures. If for some reason you omit a callback or use an incorrect callback signature, your UI customization will fail silently unless you turn on error reporting when you debug your solution.

Note that if any portion of the XML markup specified by an add-in and returned by  **GetCustomUI** does not adhere to the Office Fluent UI XML schema, none of the UI customization specified by that add-in is applied. For example, if you have a problem with one control that you have added to the ribbon, your customizations for that ribbon are not displayed.

To view any errors that your XML markup generates when it is loaded, follow these steps:


1. Click the  **File** tab, and then click **Options**.
    
2. Click  **Advanced**.
    
3. Under  **Developers**, select  **Show add-in user interface errors**.
    
4. Click  **OK** to save your changes.
    

## See also


#### Concepts


 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)<br>
 [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md)

