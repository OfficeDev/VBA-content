---
title: Application.GlobalOutlineCodes Property (Project)
ms.prod: project-server
api_name:
- Project.Application.GlobalOutlineCodes
ms.assetid: a63d1a87-5c87-a2d6-c4da-70ab9526eaae
ms.date: 06/08/2017
---


# Application.GlobalOutlineCodes Property (Project)

Gets or sets an  **[OutlineCodes](outlinecodes-object-project.md)** collection in the Global.mpt file, along with enterprise text custom fields that use a lookup table. Read/write **OutlineCodes**.


## Syntax

 _expression_. **GlobalOutlineCodes**

 _expression_ A variable that represents an **Application** object.


## Remarks




 **Note**  In Project, an outline code is any enterprise custom field that uses a single- or multi-level text lookup table. Enterprise custom fields can only be created in Project Web App, with Project Server Interface (PSI) methods, or with the client-side object model (CSOM) in Project. The enterprise global template does not store enterprise custom fields.

When Project Professional is not connected with Project Server, the  **GlobalOutlineCodes** property gets only the collection of outline codes in the Global.mpt file on the local computer. When Project Professional is connected with Project Server, the collection of outline codes includes those in the Global.mpt file plus the enterprise text custom fields with a lookup table.


## Example

The following example lists all of the outline codes in Project Server that use a text lookup table.


```vb
Sub ListGlobalOutlineCodes() 
    Dim i As Integer 
    Dim numCF_withLUTs As Integer 
    numCF_withLUTs = GlobalOutlineCodes.count 
 
    For i = 1 To numCF_withLUTs 
        Debug.Print GlobalOutlineCodes.Item(i).Name 
    Next i 
End Sub
```

In Project Server 2013 , the default text custom fields that have a lookup table include the following: 


- Cost Type
    
- Health
    
- Project Departments
    
- Resource Departments
    
- RBS
    

