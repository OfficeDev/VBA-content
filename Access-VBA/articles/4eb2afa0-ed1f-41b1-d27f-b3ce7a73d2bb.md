
# AddMenu Macro Action

 **Last modified:** July 28, 2015

 _ **Applies to:** Access 2013 | Access 2016_

This article describes the basic operation of the  **AddMenu** macro action.

You can use the  **AddMenu** action to create:

- Custom menus on the  **Add-Ins** tab for a particular form or report.
    
- A custom shortcut menu for a form, report, or control. The custom shortcut menu replaces the built-in shortcut menu for the form, report, or control.
    
- A global shortcut menu. The global shortcut menu replaces the built-in shortcut menu for fields in table and query datasheets, forms, and reports, except where you've added a custom shortcut menu for a form, report, or control.
    

## Setting

The  **AddMenu** action has the following arguments.



|**Action argument**|**Description**|
|:-----|:-----|
|**Menu Name**|The name of the menu, for example, "Report Commands" or "Tools". To create an access key so that you can use the keyboard to choose the menu, type an ampersand ( **&;** ) before the letter you want to be the access key. This letter will be underlined in the menu name on the **Add-Ins** tab.|
|**Menu Macro Name**|The name of the macro group that contains the macros for the menu's commands. This is a required argument.
 **Note**  If you run a macro containing the  **AddMenu** action in a library database, Microsoft Office Access 2007 looks for the macro group with this name in the current database only.

|
|**Status Bar Text**|The text to display in the status bar when the menu is selected. This argument is ignored for shortcut menus.|

## Remarks

To run the  **AddMenu** action in a Visual Basic for Applications (VBA) module, use the **AddMenu** method of the **DoCmd** object. You can also set the **MenuBar** or **ShortcutMenuBar** property in VBA to create a custom menu on the **Add-Ins** tab or to attach a custom shortcut menu to a form, report, or control. You can set the **ShortcutMenuBar** property of the **Application** object to create a global shortcut menu.

