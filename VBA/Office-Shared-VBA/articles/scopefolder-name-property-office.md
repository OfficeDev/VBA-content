---
title: ScopeFolder.Name Property (Office)
keywords: vbaof11.chm259001
f1_keywords:
- vbaof11.chm259001
ms.prod: office
api_name:
- Office.ScopeFolder.Name
ms.assetid: da1cc239-2988-2b57-11d1-8313ae3d5566
ms.date: 06/08/2017
---


# ScopeFolder.Name Property (Office)

Gets the name of a searchable folder. Read-only.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **ScopeFolder** object.


### Return Value

String


## Remarks

 **ScopeFolder** objects are intended for use with the **SearchFolders** collection. The **SearchFolders** collection defines the folders that are searched.


## Example

The following example displays a message box with the name of the folder that will be searched.


```
Dim sf As ScopeFolder 
 Dim strScopeFolder As String 
 
 Set sf = SearchScopes.Item(1).ScopeFolder 
 strScopeFolder = sf.Name 
 
 MsgBox ("The name of the folder that will be searched is " &amp; strScopeFolder) 

```


## See also


#### Concepts


[ScopeFolder Object](scopefolder-object-office.md)
#### Other resources


[ScopeFolder Object Members](scopefolder-members-office.md)

