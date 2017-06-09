---
title: Project.MapList Property (Project)
keywords: vbapj.chm132404
f1_keywords:
- vbapj.chm132404
ms.prod: project-server
api_name:
- Project.Project.MapList
ms.assetid: b124f86e-fec6-ab92-93ff-5db4eff16892
ms.date: 06/08/2017
---


# Project.MapList Property (Project)

Gets a  **[List](list-object-project.md)** object representing the list of data maps in the project. Read-only **List**.


## Syntax

 _expression_. **MapList**

 _expression_ A variable that represents a **Project** object.


## Example

The following example prints the list of data maps in the active project.


```vb
Sub TestMapList() 
    Dim lst As List 
    Dim numLists As Integer 
    Dim i As Integer 
 
    Set lst = ActiveProject.MapList 
    numLists = lst.Count 
 
    For i = 1 To numLists 
        Debug.Print lst.Item(i) 
    Next i 
 
End Sub
```

Following is the default map list in Project: 


- Default task information
    
- Task "Export Table" map
    
- Resource "Export Table" map
    
- Task list with embedded assignment rows
    
- Task and resource PivotTable report
    
- Top Level Tasks list
    
- "Who Does What" report
    
- Earned value information
    
- Cost data by task
    
- Compare to Baseline
    



