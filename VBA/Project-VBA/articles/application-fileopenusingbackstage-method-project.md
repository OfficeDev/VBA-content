---
title: Application.FileOpenUsingBackstage Method (Project)
keywords: vbapj.chm1010
f1_keywords:
- vbapj.chm1010
ms.prod: project-server
ms.assetid: 8e67d279-cbe6-4cfc-f809-ab83c6298e2f
ms.date: 06/08/2017
---


# Application.FileOpenUsingBackstage Method (Project)
Displays the  **Open** tab in the Backstage view.

## Syntax

 _expression_. **FileOpenUsingBackstage**

 _expression_ A variable that represents an **Application** object.


### Return value

 **Boolean**

The return value is  **True** if Project displays the **Open** tab in the Backstage view; otherwise, **False** if there is an error.


## Example

The following line of code prints  `Open in Backstage: True` in the **Immediate** window of the VBE.


```vb
Debug.Print "Open in Backstage: " &; Application.FileOpenUsingBackstage()
```


## See also


#### Concepts


[FileOpenEx Method](application-fileopenex-method-project.md)
