
# CloseWindow Macro Action

 **Last modified:** July 28, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You can use the  **CloseWindow** action to close either a specified Access document tab or the active document tab if none is specified.


## Setting

The  **CloseWindow** action has the following arguments.



|**Action argument**|**Description**|
|:-----|:-----|
|**Object Type**|The type of object whose document tab you want to close. Click  **Table**,  **Query**,  **Form**,  **Report**,  **Macro**,  **Module**,  **Data Access Page**,  **Server View**,  **Diagram**,  **Stored Procedure**, or  **Function** in the **Object Type** box in the **Action Arguments** section of the Macro Builder pane. To select the active document tab, leave this argument blank.
 **Note**  If you are closing a module in the Visual Basic Editor, you must use  **Module** in the **Object Type** argument.

|
|**Object Name**|The name of the object to be closed. The  **Object Name** box shows all objects in the database of the type selected by the **Object Type** argument. Click the object to close. If you leave the **Object Type** argument blank, leave this argument blank also.|
|**Save**|Whether to save changes to the object when it is closed. Click  **Yes** (save the object), **No** (close the object without saving it), or **Prompt** (prompt the user whether or not to save the object). The default is **Prompt**.|

## Remarks

The  **CloseWindow** action works on all database objects that the user can explicitly open or close. This action has the same effect as selecting an object and then closing it by right-clicking the object's document tab and then clicking **Close** on the shortcut menu, or clicking the **Close** button for the object.

If the  **Save** argument is set to **Prompt** and the object hasn't already been saved before the **CloseWindow** action is carried out, a dialog box prompts the user to save the object before the macro closes it. If you have set the **Warnings On** argument of the **SetWarnings** action to **No**, the dialog box is not displayed and the object is automatically saved.

To run the  **CloseWindow** action in a Visual Basic for Applications (VBA) module, use the **CloseWindow** method of the **DoCmd** object.

