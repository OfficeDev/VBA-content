
# JournalItem.AutoResolvedWinner Property (Outlook)

 **Last modified:** July 28, 2015

Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.

## Syntax

 _expression_. **AutoResolvedWinner**

 _expression_A variable that represents a  **JournalItem** object.


## Remarks

A value of  **False** does not necessarily indicate that the item is a loser of an automatic conflict resolution. The item could be in conflict with another item.

If an item has  ** [Conflicts.Count](4a7445ff-8628-50d6-f4c0-ada85f3b3f5c.md)** of its ** [JournalItem.Conflicts](27e68a60-745a-43a3-b1db-e4c80a9e4a38.md)** property greater than zero and if its **AutoResolvedWinner** property is **True**, it is a winner of an automatic conflict resolution. On the other hand, if the item is in conflict and has its  **AutoResolvedWinner** property as **False**, it is a loser in an automatic conflict resolution.


## See also


#### Concepts


 [JournalItem Object](6e850295-39f9-47b8-e866-9622e9958c69.md)
#### Other resources


 [JournalItem Object Members](13a0cd10-44bc-a167-c613-93985f698d95.md)
