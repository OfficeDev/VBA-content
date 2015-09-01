
# Items.GetLast Method (Outlook)

 **Last modified:** July 28, 2015

Returns the last object in the collection. 

## Syntax

 _expression_. **GetLast**

 _expression_A variable that represents an  **Items** object.


### Return Value

An  **Object** value that represents the last object contained by the collection.


## Remarks

It returns  **Nothing** if no last object exists, for example, if the collection is empty. To ensure correct operation of the ** [GetFirst](142a6174-118e-6256-0511-8ae9e142e555.md)**,  **GetLast**,  ** [GetNext](01c49c21-d9f9-37c4-8c64-ff8e2b1f9462.md)**, and  ** [GetPrevious](5dde47f8-2bd8-fdbe-d6e7-b1381e8a97a6.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


 [Items Object](3a99730b-e62a-5ca6-f6ec-911c95173242.md)
#### Other resources


 [Items Object Members](bcc2cf6c-b6fb-e1a2-1d5c-d7e2bdf6b7dc.md)
