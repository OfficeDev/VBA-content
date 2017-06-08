---
title: DoCmd.OpenDiagram Method (Access)
keywords: vbaac10.chm4650
f1_keywords:
- vbaac10.chm4650
ms.prod: access
api_name:
- Access.DoCmd.OpenDiagram
ms.assetid: a9736e57-eb82-77d7-c57a-8c793333392a
ms.date: 06/08/2017
---


# DoCmd.OpenDiagram Method (Access)

The  **OpenDiagram** method carries out the OpenDiagram action in Visual Basic.


## Syntax

 _expression_. **OpenDiagram**( ** _DiagramName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DiagramName_|Required|**Variant**|A string expression that's the valid name of a database diagram in the current database. If you execute Visual Basic code containing the  **OpenDiagram** method in a library database, Microsoft Access looks for the database diagram with this name first in the library database, then in the current database.|

## Remarks

In a Microsoft Access project, you can use the  **OpenDiagram** method to open a database diagram in Design view.


## Example

The following example opens the database diagram named "Data Model".


```vb
DoCmd.OpenDiagram " Data Model"
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

