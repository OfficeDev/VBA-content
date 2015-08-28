
# Application.FileBuildID Property (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets the file build identification number (ID) of the specified project. The build ID consists of the version and build of the Project application that created the file. Read-only  **String**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **FileBuildID**( **_Name_**,  **_UserID_**,  **_DatabasePassWord_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Required| **String**|The name of a project file, source file, or data source.|
|UserID|Optional| **String**|A user ID to use when accessing a database. If Name isn't a database,UserID is ignored.|
|DatabasePassWord|Optional| **Variant**|A password to use when accessing a database. If Name isn't a database,DatabasePassWord is ignored.|

## Remarks
<a name="sectionSection1"> </a>

The  **FileBuildID** property can get the file build ID of a project file without actually opening it.


## Example
<a name="sectionSection2"> </a>

The following example gets the build ID for the Test.mpp project. If the Project build that created the file is 15.0.4027.1000, the  **FileBuildID** value is "15,0,4027,1000".


```
Sub File_BuildID()
    Dim ProjID As String

    ProjID = Application.FileBuildID("C:\Project\VBA\Samples\Test.mpp")
    Debug.Print ProjID
End Sub
```

