---
title: Application.FileBuildID Property (Project)
ms.prod: project-server
api_name:
- Project.Application.FileBuildID
ms.assetid: 6fae0673-614d-6cb2-31c2-bff9eabeecc9
ms.date: 06/08/2017
---


# Application.FileBuildID Property (Project)

Gets the file build identification number (ID) of the specified project. The build ID consists of the version and build of the Project application that created the file. Read-only  **String**.


## Syntax

 _expression_. **FileBuildID**( ** _Name_**, ** _UserID_**, ** _DatabasePassWord_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of a project file, source file, or data source.|
| _UserID_|Optional|**String**|A user ID to use when accessing a database. If  _Name_ isn't a database, _UserID_ is ignored.|
| _DatabasePassWord_|Optional|**Variant**|A password to use when accessing a database. If  _Name_ isn't a database, _DatabasePassWord_ is ignored.|

## Remarks

The  **FileBuildID** property can get the file build ID of a project file without actually opening it.


## Example

The following example gets the build ID for the Test.mpp project. If the Project build that created the file is 15.0.4027.1000, the  **FileBuildID** value is "15,0,4027,1000".


```vb
Sub File_BuildID()
    Dim ProjID As String

    ProjID = Application.FileBuildID("C:\Project\VBA\Samples\Test.mpp")
    Debug.Print ProjID
End Sub
```


