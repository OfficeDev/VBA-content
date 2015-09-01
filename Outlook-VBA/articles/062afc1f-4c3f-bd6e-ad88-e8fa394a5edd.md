
# RemoteItem.Delete Method (Outlook)

 **Last modified:** July 28, 2015

Removes the item from the folder that contains the item.

## Syntax

 _expression_. **Delete**

 _expression_A variable that represents a  **RemoteItem** object.


## Remarks

The  **Delete** method deletes a single item in a collection. To delete all items in the ** [Items](441820e7-5fe8-e5ef-83c0-9c87fd3dc9e3.md)** collection of a folder, you must delete each item starting with the last item in the folder. For example, in the items collection of a folder, `AllItems`, if there are  `n` number of items in the folder, start deleting the item at `AllItems.Item(n)`, decrementing the index each time until you delete  `AllItems.Item(1)`.

The  **Delete** method moves the item from the containing folder to the **Deleted Items** folder. If the containing folder is the **Deleted Items** folder, the **Delete** method removes the item permanently.


## See also


#### Concepts


 [RemoteItemObject](6302aaff-cdcf-4d86-60f1-4bed15540d9f.md)
 [Delete All Items and Subfolders in the Deleted Items Folder](359a416b-43d4-396e-e348-5624c4ca3599.md)
#### Other resources


 [RemoteItemMembers](15c0872e-88cc-9b9b-c31e-c15d6971e6e0.md)
