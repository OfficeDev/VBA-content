---
title: WebNavigationBarSet.DeleteSetAndInstances Method (Publisher)
keywords: vbapb10.chm8519683
f1_keywords:
- vbapb10.chm8519683
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.DeleteSetAndInstances
ms.assetid: 89bbd9b9-d0c9-ecac-eb3e-7425bd177aec
ms.date: 06/08/2017
---


# WebNavigationBarSet.DeleteSetAndInstances Method (Publisher)

Deletes a Web navigation bar set and all instances of it in the current document.


## Syntax

 _expression_. **DeleteSetAndInstances**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


## Example

The following example iterates through the  **WebNavigationBarSets** collection and deletes each set from the active document.


```vb
Dim objWebNavBarSet As WebNavigationBarSet 
For Each objWebNavBarSet In ActiveDocument.WebNavigationBarSets 
 objWebNavBarSet.DeleteSetAndInstances 
Next objWebNavBarSet
```


