
# Connections.Count Property (DAO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

Returns the number of  **[Connection](f469b04e-2539-6b53-31f2-85fe22fcc2fc.md)** objects in the **[Connections](65d073be-a84b-e3f2-cb43-b87ffa60e497.md)** collection.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Connections** object.


## Remarks

Because members of a collection begin with 0, you should always code loops starting with the 0 member and ending with the value of the  **Count** property minus 1. If you want to loop through the members of a collection without checking the **Count** property, you can use a **For Each...Next** command.

The  **Count** property setting is never Null. If its value is 0, there are no objects in the collection.

