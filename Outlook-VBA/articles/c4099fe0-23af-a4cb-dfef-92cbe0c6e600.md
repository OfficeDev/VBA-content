
# TaskItem.Categories Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.

## Syntax

 _expression_. **Categories**

 _expression_A variable that represents a  **TaskItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under  **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


#### Concepts


 [TaskItem Object](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)
#### Other resources


 [TaskItem Object Members](97234a76-2fc5-bbe4-2e14-25ae18694fc9.md)
