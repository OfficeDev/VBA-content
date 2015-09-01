
# TaskRequestItem.Categories Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.

## Syntax

 _expression_. **Categories**

 _expression_A variable that represents a  **TaskRequestItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under  **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


#### Concepts


 [TaskRequestItem Object](2908a28a-634c-e786-aa53-f3e32038b727.md)
#### Other resources


 [TaskRequestItem Object Members](d43114ee-be91-ff02-3424-525da2cf3a50.md)
