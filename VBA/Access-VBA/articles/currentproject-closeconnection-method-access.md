---
title: CurrentProject.CloseConnection Method (Access)
keywords: vbaac10.chm12716
f1_keywords:
- vbaac10.chm12716
ms.prod: access
api_name:
- Access.CurrentProject.CloseConnection
ms.assetid: f2feac44-e509-48d7-e815-e0cf2935d7b9
ms.date: 06/08/2017
---


# CurrentProject.CloseConnection Method (Access)

You can use the  **CloseConnection** method to close the current connection between the **CurrentProject** object in a Microsoft Access project (.adp) or Access database and the database specified in the project's base connection string.


## Syntax

 _expression_. **CloseConnection**

 _expression_ A variable that represents a **CurrentProject** object.


### Return Value

Nothing


## Remarks

The  **CloseConnection** method closes the current connection of the Access project, database, or data source control, frees the ADO **Connection** object, and sets the **Connection** property to **Null**. The **BaseConnectionString** property is left unchanged. Users are prevented from calling _datasourcecontrol_.Connection.Close and must use this method instead.

The  **CloseConnection** method is useful when you have opened a Microsoft Access database from another application through Automation.


## See also


#### Concepts


[CurrentProject Object](currentproject-object-access.md)

