---
title: Pane.Index Property (Project)
ms.prod: project-server
api_name:
- Project.Pane.Index
ms.assetid: 6989c013-eb83-05ea-77c4-1c90517f389b
ms.date: 06/08/2017
---


# Pane.Index Property (Project)

Gets the index of a  **Pane** object in the containing object. Read-only **Variant**.


## Syntax

 _expression_. **Index**

 _expression_ A variable that represents a **Pane** object.


## Remarks

A  **Pane** object can be accessed through the **Window** object in either a **Windows** or **Windows2** collection. For example, `windows2.Item(1).TopPane.Index` has the value 1.

The  **Index** properties of other objects are used in similar ways. For an example, see the **[Index](project-index-property-project.md)** property of the **Project** object.


