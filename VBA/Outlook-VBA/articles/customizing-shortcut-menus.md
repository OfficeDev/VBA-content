---
title: Customizing Shortcut Menus
ms.prod: outlook
ms.assetid: ed6a98a3-243b-80ee-51ae-57dba6d8715a
ms.date: 06/08/2017
---


# Customizing Shortcut Menus

You can customize several different shortcut menus in Microsoft Outlook by using your add-in to change, disable, or remove existing menu items, or to add new menu items.

You customize shortcut menus by using Microsoft Office Fluent user interface (UI) extensibility, just as you would to customize the user interface on a ribbon in an explorer or inspector. 

Because  [CommandBar](http://msdn.microsoft.com/library/78603954-40aa-64cb-c407-2e0820d65231%28Office.15%29.aspx) objects have been deprecated since Outlook, shortcut menu events of the [Application](application-object-outlook.md) object that relied on the **CommandBar** object are being deprecated as well, and might not work as expected. These events include the following:


-  **AttachmentContextMenuDisplay** event
    
-  **ContextMenuClose** event
    
-  **FolderContextMenuDisplay** event
    
-  **ItemContextMenuDisplay** event
    
-  **ShortcutContextMenuDisplay** event
    
-  **StoreContextMenuDisplay** event
    
-  **ViewContextMenuDisplay** event
    

 To customize shortcut menus, implement the **[IRibbonExtensibility](http://msdn.microsoft.com/library/b27a7576-b6f5-031e-e307-78ef5f8507e0%28Office.15%29.aspx)** interface in your add-in. Specifically, implement the **[GetCustomUI](http://msdn.microsoft.com/library/a0106415-999e-94da-379c-70fb7aa6119f%28Office.15%29.aspx)** method of the **IRibbonExtensibility** interface so that when Office calls the **GetCustomUI** method and specifies **Microsoft.Outlook.Explorer** as the ribbon ID, the method loads the custom shortcut menu that is delimited by the `contextMenus` tag in XML. For a complete listing of ribbon identifiers, see [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md).

## Customizable Shortcut Menus

You can use Office Fluent UI extensibility to customize the following types of shortcut menus:


- Alternative interactions shortcut menus.
    
- Attachment shortcut menus.
    
- Folder shortcut menus.
    
- Item, flagged item, new item, and item selection shortcut menus.
    
- Persona shortcut menus.
    
- A shortcut menu for a shortcut in the  **Shortcuts** module.
    
- Store shortcut menus.
    
- View and view user interface shortcut menus.
    
For more information about customizing shortcut menus, including examples, see  [Extending the User Interface in Outlook 2010](http://msdn.microsoft.com/en-us/library/ee692172%28office.14%29.aspx) on the MSDN Web site.


## See also


#### Concepts


 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)

