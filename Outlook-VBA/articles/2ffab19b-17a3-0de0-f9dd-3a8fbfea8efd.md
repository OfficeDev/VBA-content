
# DistListItem.GetInspector Property (Outlook)

 **Last modified:** July 28, 2015

Returns an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)**object that represents an inspector initialized to contain the specified item. Read-only.

## Syntax

 _expression_. **GetInspector**

 _expression_A variable that represents a  **DistListItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the ** [Application.ActiveInspector](3f2b6491-7b4b-8165-327e-b319711d5656.md)**method and setting the  ** [Inspector.CurrentItem](eaaf0192-a169-c107-95a6-b8e759a3b873.md)**property. If an  **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


 [DistListItem Object](027c3986-abff-d9b1-ecc2-26d60805e952.md)
#### Other resources


 [DistListItem Object Members](3ba4af84-ce84-61d9-1bc9-fab41bf6f125.md)
