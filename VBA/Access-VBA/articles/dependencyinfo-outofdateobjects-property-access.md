---
title: DependencyInfo.OutOfDateObjects Property (Access)
keywords: vbaac10.chm13276
f1_keywords:
- vbaac10.chm13276
ms.prod: access
api_name:
- Access.DependencyInfo.OutOfDateObjects
ms.assetid: 3e6465c0-c1e4-0b26-de2e-0610e3a40273
ms.date: 06/08/2017
---


# DependencyInfo.OutOfDateObjects Property (Access)

Returns a  **[DependencyObjects](dependencyobjects-object-access.md)** collection that represents the **[AccessObject](accessobject-object-access.md)** objects for which the dependency information is outdated. Read-only.


## Syntax

 _expression_. **OutOfDateObjects**

 _expression_ A variable that represents a **DependencyInfo** object.


## Remarks

You can use the following code to update the dependency information for all of the objects in the database:


```vb
Application.CurrentProject.UpdateDependencyInfo
```


## See also


#### Concepts


[DependencyInfo Object](dependencyinfo-object-access.md)

