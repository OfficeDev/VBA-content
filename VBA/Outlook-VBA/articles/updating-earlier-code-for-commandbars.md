---
title: Updating Earlier Code for CommandBars
ms.prod: outlook
ms.assetid: 58bc6957-fa1e-72ac-1836-a2a411e089c2
ms.date: 06/08/2017
---


# Updating Earlier Code for CommandBars

 In versions prior to Outlook, Outlook implemented the following items in the user interface as command bars:


- Menu bars, toolbars, and shortcut menus.
    
- Menus on menu bars and toolbars.
    
- Submenus on menus, submenus, and shortcut menus.
    



Command bars have been deprecated since Outlook 2010. Explorer and inspector windows no longer use menu bars and toolbars. Instead, they use the Microsoft Office Fluent ribbon. Programmatically, although your add-in or VBScript code that customized command bars in an explorer or inspector might still work in Outlook, those who use your solution might not easily find the customizations that appear on the  **Add-ins** tab of the customized ribbon in the explorer or inspector.


**Note** To find out more about issues to consider before you upgrade existing solutions, read [Developer Issues When Upgrading Solutions to Outlook 2010](https://msdn.microsoft.com/en-us/library/office/ff864759(v=office.14).aspx).

Instead of using the  **CommandBars** property of the [Explorer](explorer-object-outlook.md) and [Inspector](inspector-object-outlook.md) objects, use ribbon extensibility to customize the ribbon and to customize any menus and submenus off the ribbon. Ribbon extensibility requires an add-in that implements the [IRibbonExtensibility](http://msdn.microsoft.com/library/b27a7576-b6f5-031e-e307-78ef5f8507e0%28Office.15%29.aspx) interface. 

For more information about customizing the ribbon in Outlook, see [Overview of Customizing the Ribbon](overview-of-customizing-the-ribbon.md).
Consistent with the deprecation of command bars in the explorer and inspector windows, do not rely on the  [CommandBar](http://msdn.microsoft.com/library/78603954-40aa-64cb-c407-2e0820d65231%28Office.15%29.aspx) object for your custom menus; instead, use an add-in through the **IRibbonExtensibility** interface to extend them. For more information, see [Customizing Shortcut Menus](customizing-shortcut-menus.md).

## See also


#### Concepts


 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)

