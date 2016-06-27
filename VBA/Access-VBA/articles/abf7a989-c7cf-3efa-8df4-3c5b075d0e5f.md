
# LockNavigationPane Macro Action

 **Last modified:** July 28, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You can use the  **LockNavigationPane** action to prevent users from deleting database objects that are displayed in the Navigation Pane.


## Setting

The  **LockNavigationPane** action has the following argument.



|**Action argument**|**Description**|
|:-----|:-----|
|**Lock**|Select  **Yes** to lock the Navigation Pane, or **No** to unlock the Navigation Pane.|

## Remarks

Locking the Navigation Pane prevents you from deleting database objects or cutting database objects to the clipboard. It does  _not_ prevent you from performing any of the following operations:


- Copying database objects to the clipboard
    
- Pasting database objects from the clipboard
    
- Displaying or hiding the Navigation Pane
    
- Selecting different Navigation Pane organization schemes
    
- Showing or hiding sections of the Navigation Pane
    
To run the  **LockNavigationPane** action in a VBA module, use the **LockNavigationPane** method of the **DoCmd** object.

