
# Conversation.GetRootItems Method (Outlook)

 **Last modified:** July 28, 2015

Returns a  ** [SimpleItems](b929ae28-fe5f-607e-37b5-ed6a304d4896.md)** collection that contains all root items in the conversation.

## Syntax

 _expression_. **GetRootItems**

 _expression_A variable that represents a  ** [Conversation](2705d38a-ebc0-e5a7-208b-ffe1f5446b1b.md)** object.


### Return Value

A  **SimpleItems** collection that includes the root item or all root items of the conversation.


## Remarks

A conversation can have one or more root items. For example, if the root item of the conversation has three child items and the root item is permanently deleted, all three child items become root items.

If all items are deleted from the conversation after the  ** [Conversation](2705d38a-ebc0-e5a7-208b-ffe1f5446b1b.md)** object has been obtained, **GetRootItems** returns a **SimpleItems** collection with zero objects. In this case, the ** [Count](2656676b-ee82-aad0-21b9-8ca963cb57d2.md)** property of the **SimpleItems** collection returns 0.


## See also


#### Concepts


 [Conversation Object](2705d38a-ebc0-e5a7-208b-ffe1f5446b1b.md)
#### Other resources


 [Conversation Object Members](09ff1e8e-7c5a-0b1e-e8e2-e259f66f71c8.md)
