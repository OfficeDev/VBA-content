
# RunMenuCommand Macro Action

 **Last modified:** July 28, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You can use the  **RunMenuCommand** action to run a built-in Microsoft Access command.


## Setting

The  **RunMenuCommand** action has the following action argument.



|**Action argument**|**Description**|
|:-----|:-----|
|**Command**|The name of the command you want to run. The  **Command** box shows the available built-in commands in Access, in alphabetical order. This is a required argument.|

## Remarks

You can use the  **RunMenuCommand** action to run an Access command from a custom menu bar, global menu bar, custom shortcut menu, or global shortcut menu.

You can use the  **RunMenuCommand** action in a macro with conditional expressions to run a command depending on certain conditions.


 **Note**  Clicking the  **File** tab and then clicking **Recent** shows the most recently used databases. You can click one of these databases instead of clicking **Open**. These database items don't appear in the drop-down list box for the  **Command** argument, and aren't available by using the **RunMenuCommand** action in a macro.

When you convert an Access database from a previous version of Access, some commands may no longer be available. A command may have been renamed, moved to a different menu, or may no longer be available in Access. The  **DoMenuItem** actions for such commands can't be converted to **RunMenuCommand** actions. When you open the macro, Access will display a **RunMenuCommand** action with a blank **Command** argument for such commands. You must edit the macro and enter a valid command argument, or delete the **RunMenuCommand** action.

To run the  **RunMenuCommand** action in a Visual Basic for Applications (VBA) module, use the **RunCommand** method of the **Application** object. (This is equivalent to the **RunCommand** method of the **DoCmd** object.)

