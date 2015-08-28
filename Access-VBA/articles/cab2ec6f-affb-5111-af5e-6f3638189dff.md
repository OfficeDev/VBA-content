
# ListBox.AllowValueListEdits Property (Access)

 **Last modified:** July 28, 2015

Gets or sets whether the  **Edit List Items** command is available when the user right-clicks a list box. Read/write **Boolean**.

## Syntax

 _expression_. **AllowValueListEdits**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

The  **AllowValueEditLists** property determines whether the **Edit List Items** command is available when the user right-clicks a list box that's bound to a Lookup field.

If the Lookup field is bound to a list of values, then the  **Edit List Items** dialog box is displayed when the user clicks **Edit List Items**. The user can then add, delete, or edit the items to be displayed in the list box.

If the Lookup field is bound to a table or query, then the form specified by the  **ListItemsEditForm** property is diplayed when the user clicks **Edit List Items**. The user can use the form to add, delete, or edit the items to be displayed in the list box.

The  **AllowValueEditLists** property is not available for list boxes on a report.


## See also


#### Concepts


 [ListBox Object](6bc00755-34e7-4fc2-8e72-40dae2010dd8.md)
#### Other resources


 [ListBox Object Members](d87ad51b-9a46-21f3-f6d6-ef98ea8aaf6d.md)
