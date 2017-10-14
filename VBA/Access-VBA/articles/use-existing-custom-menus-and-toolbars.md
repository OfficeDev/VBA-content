---
title: Use Existing Custom Menus and Toolbars
ms.prod: access
ms.assetid: 581167ca-7c9c-4402-a9b7-49393750079c
ms.date: 06/08/2017
---


# Use Existing Custom Menus and Toolbars

This topic describes how custom toolbars and menu bars that you created in earlier versions of Access behave when you open those older databases in Access. This topic also explains how to turn off the ribbon so that you can use custom toolbars and menu bars.


## How earlier version databases behave in Access 2010

If you have an .mdb file that contains custom toolbars, custom menu bars, or a custom startup behavior, those options apply when you open that earlier version database in Access 2010 and when you convert the earlier version database to an Access 2010 file (an .accdb or .accde file). For example, if you turned off built-in toolbars and replaced the default menu bar in a database created in Access 2003, that behavior remains in place when you open the database in Access 2010.

In addition, you can hide the ribbon when you open legacy files (.mdb, .mde, and .mda files) in Access 2010, and when you use to create a legacy file — something you cannot do with .accdb and .accde files. You can define different behaviors for legacy files, because the earlier version database formats use a different working model than the new .accdb and .accde files. Earlier versions of Access opened each object in a separate window; in contrast, (by default) opens all objects in a single, tabbed document and separates open objects with tabs.

 In addition, you can turn off the ribbon for a legacy database by setting options for that database in Access 2003, or by setting global program options in . A procedure, later in this topic, explains how to perform both tasks.

When you open a legacy database and you choose to display the ribbon, any custom toolbars appear as groups on the  **Add-Ins** tab. Each group in the tab corresponds to a custom toolbar, and each group uses the name assigned to the original toolbar. However, the toolbars must be visible in the legacy database or they do not appear on the tab.

 **Reminder** To bypass custom startup behaviors, press and hold SHIFT while you open the database.

The following procedure describes how to open and use a database that contains one or more custom toolbars, how to open a database that uses customized startup behavior, and how to hide the ribbon.


## Opening and using an earlier version database that contains custom toolbars


1. Click the  **File** tab, and then click **Open**.
    
    The  **Open** dialog box appears.
    
2. Use the  **Look in** list to locate your legacy database (a .mdb or .mde file), and then click **Open**. Access 2010 opens the earlier version database. The database objects — the tables, forms, reports, and so on — appear in the Navigation Pane. If you set a form, switchboard, or other object to appear on startup, that object also appears in the Navigation Pane. Also, if you created custom toolbars or menu bars, they appear in the  **Add-Ins** tab as one or more groups. Each group uses the name originally assigned to the custom toolbar or menu bar.
    
3. Click the  **Add-Ins** tab. Your custom toolbars appear as one or more groups, and you can use them when doing so is logical. For example, suppose that one of your custom toolbars contains the **Print Relationships** command. Access does not enable that command until you display the relationships for the open database.
    

 **Note**  If your database does not contain a custom toolbar, the  **Add-Ins** tab remains hidden.


## Opening and using an earlier version database with custom startup behavior


 **Note**  These steps assume that you have a database created in an earlier version of Access, and that database uses custom startup settings. If not, you can ignore these steps.

 **Open a database**


1. Click the  **File** tab, and then click **Open**. The  **Open** dialog box appears.
    
2. Use the  **Look in** list to locate and open the earlier version database. opens the database and runs any startup settings. For example, if the earlier version database was set to run a parameter query before opening any data-entry forms, the dialog boxes for that query appear in Access 2010.
    

 **Note**  If the database uses Visual Basic for Applications (VBA) code, Access blocks the code by default.


## Turn off the Ribbon and use just your custom menu bars

The following procedure describes how to hide the ribbon by changing settings in Access 2003 and in Access 2010. To follow these steps, you must have a database created in an earlier version of Access, and that database must contain a custom menu bar. For more information about creating a custom menu bar, see the Help for your earlier version of Access.

 **Set Access 2003 to use a custom menu bar**


1. Using Access 2003, open your legacy database.
    
2. On the  **Tools** menu, click **Startup**. The  **Startup** dialog box appears.
    
3. From the  **Menu Bar** list, select your custom menu bar.
    
     **Note**  You must select a menu bar. You cannot select a toolbar.
4. Clear the  **Allow Built-in Toolbars** check box, click **OK**, and then close the database. When you open the database in Access 2010, Access shows the Message Bar (if necessary), the custom menu bar set for the database, and any other startup settings, such as a form and any custom toolbars.
    
 **Set Access 2010 to use custom menu bars**


1. Open your legacy database in Access 2010.
    
2. Click the  **File** tab, and then click **Options**.
    
3. In the  **Access Options** dialog box, click **Current Database**.
    
4. Under  **Ribbon and Toolbar Options**, clear the  **Allow Full Menus** and **Allow Built-in Toolbars** check boxes.
    
5. Click  **OK**.
    

