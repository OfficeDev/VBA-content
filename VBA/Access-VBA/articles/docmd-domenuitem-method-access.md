---
title: DoCmd.DoMenuItem Method (Access)
keywords: vbaac10.chm4148
f1_keywords:
- vbaac10.chm4148
ms.prod: access
api_name:
- Access.DoCmd.DoMenuItem
ms.assetid: b897bfdb-7f03-2b42-2bfd-219a2f4aa21b
ms.date: 06/08/2017
---


# DoCmd.DoMenuItem Method (Access)

Displays the appropriate menu or toolbar command for Microsoft Access.


## Syntax

 _expression_. **DoMenuItem**( ** _MenuBar_**, ** _MenuName_**, ** _Command_**, ** _Subcommand_**, ** _Version_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MenuBar_|Required|**Variant**|Use the intrinsic constant **acFormBar** for the menu bar in Form view. For other views, use the number of the view in the menu bar argument list, as shown in the Macro window in previous versions of Microsoft Access (count down the list, starting from 0).|
| _MenuName_|Required|**Variant**|You can use one of the following intrinsic constants.<ul><li><p><b>acFile</b></p></li><li><p><b>acEditMenu</b></p></li><li><p><b>acRecordsMenu</b></p></li><li><p>You can use <b>acRecordsMenu</b>  only for the Form view menu bar in Microsoft Access version 2.0 and Microsoft Access 95 databases. For other menus, use the number of the menu in the menu name argument list, as shown in the Macro window in previous versions of Microsoft Access (count down the list, starting from 0).</p></li></ul>|
| _Command_|Required|**Variant**|You can use one of the following intrinsic constants.<ul><li><p><b>acNew</b></p></li><li><p><b>acSaveForm</b></p></li><li><p><b>acSaveFormAs</b></p></li><li><p><b>acSaveRecord</b></p></li><li><p><b>acUndo</b></p></li><li><p><b>acCut</b></p></li><li><p><b>acCopy</b></p></li><li><p><b>acPaste</b></p></li><li><p><b>acDelete</b></p></li><li><p><b>acSelectRecord</b></p></li><li><p><b>acSelectAllRecords</b></p></li><li><p><b>acObjacRefreshect</b></p></li><li><p>For other commands, use the number of the command in the command argument list, as shown in the Macro window in previous versions of Microsoft Access (count down the list, starting from 0).</p></li></ul>|
| _Subcommand_|Optional|**Variant**|You can use one of the following intrinsic constants.<ul><li><p><b>acObjectVerb</b></p></li><li><p><b>acObjectUpdate</b></p></li><li><p>The <b>acObjectVerb</b>  constant represents the first command on the submenu of the <span class="ui">Object</span> command on the <span class="ui">Edit</span> menu. The type of object determines the first command on the submenu. For example, this command is Edit for a Paintbrush object that can be edited.</p></li><li><p>For other commands on submenus, use the number of the subcommand in the subcommand argument list, as shown in the Macro window in previous versions of Microsoft Access (count down the list, starting from 0).</p></li></ul>|
| _Version_|Optional|**Variant**|Use the intrinsic constant  **acMenuVer70** for code written for Microsoft Access 95 databases, the intrinsic constant **acMenuVer20** for code written for Microsoft Access version 2.0 databases, and the intrinsic constant **acMenuVer1X** for code written for Microsoft Access version 1.x databases. This argument is available only in Visual Basic.<table><tr><th>**Note**</th></tr><tr><td>The default for this argument is **acMenuVer1X**, so that any code written for Microsoft Access version 1.x databases will run unchanged. If you're writing code for a Microsoft Access 95 or version 2.0 database and want to use the Microsoft Access 95 or version 2.0 menu commands with the **DoMenuItem** method, you must set this argument to **acMenuVer70** or **acMenuVer20**.</td></tr></table>Also, when you are counting down the lists for the Menu Bar, Menu Name, Command, and Subcommand action arguments in the Macro window to get the numbers to use for the arguments in the  **DoMenuItem** method, you must use the Microsoft Access 95 lists if the Version argument is **acMenuVer70**, the Microsoft Access version 2.0 lists if the Version argument is Version, and the Microsoft Access version 1.x lists if Version is **acMenuVer1X** (or blank).<table><tr><th>**Note**</th></tr><tr><td>There is no  **acMenuVer80** setting for this argument. You can't use the **DoMenuItem** method to display Access commands (although existing **DoMenuItem** methods in Visual Basic code will still work). Use the **RunCommand** method instead.</td></tr></table>|

## Remarks

|**Note**|
|:-----|  
|In Microsoft Access 97 and later, the  **DoMenuItem** method was replaced by the **[RunCommand](application-runcommand-method-access.md)** method. The **DoMenuItem** method is included in this version of Microsoft Access only for compatibility with previous versions. When you run existing Visual Basic code containing a **DoMenuItem** method, Microsoft Access will display the appropriate menu or toolbar command for Microsoft Access 2000. However, unlike the DoMenuItem action in a macro, a **DoMenuItem** method in Visual Basic code isn't converted to a **RunCommand** method when you convert a database created in a previous version of Microsoft Access.|

Some commands from previous versions of Microsoft Access aren't available in Access, and  **DoMenuItem** methods that run these commands will cause an error when they're executed in Visual Basic. You must edit your Visual Basic code to replace or delete occurrences of such **DoMenuItem** methods.

The selections in the lists for the menu name, command, and subcommand action arguments in the Macro window depend on what you've selected for the previous arguments. You must use numbers or intrinsic constants that are appropriate for each MenuBar, MenuName, Command, and Subcommand argument.

If you leave the Subcommand argument blank but specify the Version argument, you must include the Subcommand argument's comma. If you leave the Subcommand and Version arguments blank, don't use a comma following the Command argument.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

