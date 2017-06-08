---
title: AccessObject.DateCreated Property (Access)
keywords: vbaac10.chm12752
f1_keywords:
- vbaac10.chm12752
ms.prod: access
api_name:
- Access.AccessObject.DateCreated
ms.assetid: 68a6fd13-2831-386f-0328-274e43219578
ms.date: 06/08/2017
---


# AccessObject.DateCreated Property (Access)

Returns a  **Date** indicating the date and time when the design of the specified object was last modified. Read-only.


## Syntax

 _expression_. **DateCreated**

 _expression_ A variable that represents an **AccessObject** object.


## Example

The following example lists all the reports in the current database and when their designs were created and modified.


```vb
Dim acobjLoop As AccessObject 
 
For Each acobjLoop In CurrentProject.AllReports 
 With acobjLoop 
 Debug.Print .Name &; " - Created " &; .DateCreated _ 
 &; " - Modified " &; .DateModified 
 End With 
Next acobjLoop
```


## See also


#### Concepts


[AccessObject Object](accessobject-object-access.md)

