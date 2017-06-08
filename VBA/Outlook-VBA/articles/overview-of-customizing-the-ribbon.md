---
title: Overview of Customizing the Ribbon
ms.prod: outlook
ms.assetid: ee49751d-9eae-357c-5fa9-0b2dd4ff0890
ms.date: 06/08/2017
---


# Overview of Customizing the Ribbon

Similar to other Microsoft Office applications such as Microsoft Word, Microsoft Excel, and Microsoft PowerPoint, Microsoft Outlook uses the Microsoft Office Fluent user interface ribbon in its explorer and inspector windows. In an item inspector, such as an e-mail message in compose mode, Outlook uses the ribbon to expose commands in item-specific elements that make it easy for users to identify the commands they need to complete their tasks.

To customize the ribbon programmatically, Outlook uses ribbon extensibility. Each Outlook add-in can specify a custom user interface in an XML markup file, and then implement the  **[IRibbonExtensibility](http://msdn.microsoft.com/library/b27a7576-b6f5-031e-e307-78ef5f8507e0%28Office.15%29.aspx)** interface. Office calls the **[IRibbonExtensibility.GetCustomUI](http://msdn.microsoft.com/library/a0106415-999e-94da-379c-70fb7aa6119f%28Office.15%29.aspx)** method before the **ThisAddin.Startup** method to load ribbon customizations for the explorer ribbon, and calls the **GetCustomUI** method the first time that it displays a particular type of inspector. When it is called, the **GetCustomID** method takes a ribbon ID as an argument and loads the corresponding XML that your add-in associates with that ribbon ID. Consider using a `Switch` statement when you implement the **GetCustomID** method to load the ribbon XML for various ribbons; it is probably the most efficient way to accommodate the variety of ribbons that you might customize.

For a complete listing of ribbon identifiers, see  [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md).

For a detailed discussion of the ribbon and ribbon extensibility, see  [Overview of the Office Fluent Ribbon](http://msdn.microsoft.com/library/773c202c-f5f9-c4f6-f833-0dd56eb21a8f%28Office.15%29.aspx).

## See also


#### Concepts


 [Detecting Errors](detecting-errors.md)<br>
 [Updating Earlier Code for CommandBars](updating-earlier-code-for-commandbars.md)<br>
 [Overview of the IRibbonUI Object](overview-of-the-iribbonui-object.md)<br>
 [Overview of the IRibbonControl Object](overview-of-the-iribboncontrol-object.md)<br>
 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)

