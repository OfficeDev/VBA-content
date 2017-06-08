---
title: Application.EnumDirectories Method (Visio)
keywords: vis_sdr.chm10016255
f1_keywords:
- vis_sdr.chm10016255
ms.prod: visio
api_name:
- Visio.Application.EnumDirectories
ms.assetid: 71ed7f7f-3428-5c50-2ab9-5452188dcfe0
ms.date: 06/08/2017
---


# Application.EnumDirectories Method (Visio)

Returns an array naming the folders Microsoft Visio would search, given a list of paths.


## Syntax

 _expression_ . **EnumDirectories**( **_PathsString_** , **_NameArray()_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PathsString_|Required| **String**|A string of full or partial paths separated by semicolons.|
| _NameArray()_|Required| **String**|Out parameter. An array that receives the enumerated folder names.|

### Return Value

Nothing


## Remarks

Several Visio properties such as  **AddonPaths** and **TemplatePaths** accept and receive a string interpreted to be a list of path (folder) names separated by semicolons. When the application looks for items in the named paths, it looks in the folders and all their subfolders.

The purpose of the  **EnumDirectories** method is to accept a string such as one that the **AddonPaths** property might produce and return a list of the folders that the application enumerates when processing such a string.

If the  **EnumDirectories** method succeeds, _NameArray()_ returns a one-dimensional array of _n_ strings indexed from 0 to _n_ - 1. Each string is the fully qualified name of a folder that exists. The list names those folders designated in the path list that exist and all their subfolders.

The  _NameArray()_ parameter is an out parameter that is allocated by the **EnumDirectories** method, and ownership is passed back to the caller. The caller should eventually perform the **SafeArrayDestroy** procedure on the returned array. (Microsoft Visual Basic and Visual Basic for Applications automatically free the strings referenced by the array's entries.)


## Example

The following example shows how to use the  **EnumDirectories** method to print (in the Immediate window) a list of all the folders Visio searches for add-ons.


```vb
 
Public Sub EnumDirectories_Example() 
 
 Dim strDirectoryNames() As String 
 Dim intLowerBound As Integer 
 Dim intUpperBound As Integer 
 
 Application.EnumDirectories Application.AddonPaths, strDirectoryNames 
 
 intLowerBound = LBound(strDirectoryNames) 
 intUpperBound = UBound(strDirectoryNames) 
 
 While intLowerBound <= intUpperBound 
 Debug.Print strDirectoryNames(intLowerBound) 
 intLowerBound = intLowerBound + 1 
 Wend 
 
End Sub
```


