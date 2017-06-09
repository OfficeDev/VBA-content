---
title: ScopeFolder.Path Property (Office)
keywords: vbaof11.chm259002
f1_keywords:
- vbaof11.chm259002
ms.prod: office
api_name:
- Office.ScopeFolder.Path
ms.assetid: 744d1d00-ecbf-b0ad-e119-8f4267f5f7c5
ms.date: 06/08/2017
---


# ScopeFolder.Path Property (Office)

Gets a  **String** indicating the full path of a **ScopeFolder** object. Read-only.


## Syntax

 _expression_. **Path**

 _expression_ A variable that represents a **ScopeFolder** object.


## Example

The following example displays the root path of each directory in My Computer. To retrieve this information, the example first gets the  **ScopeFolder** object at the root of My Computer. The path of this **ScopeFolder** will always be "*". As with all **ScopeFolder** objects, the root object contains a **ScopeFolders** collection. This example loops through this **ScopeFolders** collection and displays the path of each **ScopeFolder** object in it. The paths of these **ScopeFolder** objects will be "A:\", "C:\", etc.


```
Sub DisplayRootScopeFolders() 
 
 'Declare variables that reference a 
 'SearchScope and a ScopeFolder object. 
 Dim ss As SearchScope 
 Dim sf As ScopeFolder 
 
 'Loop through the SearchScopes collection 
 'and display all of the root ScopeFolders collections in 
 'the My Computer scope. 
 For Each ss In SearchScopes 
 Select Case ss.Type 
 Case msoSearchInMyComputer 
 
 'Loop through each ScopeFolder collections in 
 'the ScopeFolders collection of the 
 'SearchScope object and display the path. 
 For Each sf In ss.ScopeFolder.ScopeFolders 
 MsgBox "Path: " &amp; sf.Path 
 Next sf 
 
 Case Else 
 End Select 
 Next ss 
 
End Sub
```


## See also


#### Concepts


[ScopeFolder Object](scopefolder-object-office.md)
#### Other resources


[ScopeFolder Object Members](scopefolder-members-office.md)

