---
title: Application.AppRestore Method (Project)
keywords: vbapj.chm2011
f1_keywords:
- vbapj.chm2011
ms.prod: project-server
api_name:
- Project.Application.AppRestore
ms.assetid: f50a1158-83d1-e38e-65e6-cdc456f14bc7
ms.date: 06/08/2017
---


# Application.AppRestore Method (Project)

Restores the main window to its previous nonminimized or nonmaximized state.


## Syntax

 _expression_. **AppRestore**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Example

The following example minimizes the application and then restores its original state.


```vb
Sub RestoreApplication() 
 'Minimize the app. 
 AppMinimize 
 'Restore the app. 
 AppRestore 
End Sub
```


