
# Workspace.Close Method (DAO)

 **Last modified:** June 30, 2011

 _ **Applies to:** Access 2013 | Access 2016_

Closes an open  **Workspace**.


## Syntax

 _expression_. **Close**

 _expression_ A variable that represents a **Workspace** object.


## Remarks

If the  **Workspace** object is already closed when you use **Close**, a run-time error occurs.

An alternative to the  **Close** method is to set the value of an object variable to **Nothing** ( `Set dbsTemp = Nothing`).

