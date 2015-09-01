
# Conflicts.GetLast Method (Outlook)

 **Last modified:** July 28, 2015

Returns the last object in the  ** [Conflicts](c4e1c060-519a-a6d1-8fb2-c7dfa1e3e66f.md)** collection.

## Syntax

 _expression_. **GetLast**

 _expression_A variable that represents a  **Conflicts** object.


### Return Value

A  ** [Conflict](a7c8f12a-08ba-9fff-60b8-a02d1c7f6f33.md)** object that represents the last object contained by the collection.


## Remarks

 It returns **Nothing** if no last object exists, for example, if the collection is empty. To ensure correct operation of the ** [GetFirst](f257a9f1-d9ec-c13a-62f7-0228d55342da.md)**,  **GetLast**,  ** [GetNext](2e21ea88-c732-17ee-cd87-698fee992269.md)**, and  ** [GetPrevious](23b5d75a-e1eb-7164-df92-71e37a1ec79f.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


 [Conflicts Object](c4e1c060-519a-a6d1-8fb2-c7dfa1e3e66f.md)
#### Other resources


 [Conflicts Object Members](dcc61922-d119-1bb9-c175-a80a73599559.md)
