
# Project.ResourceGroupList Property (Project)

 **Last modified:** July 28, 2015

Gets a  ** [List](3934c2e8-d810-6571-9a33-1d41edbab87a.md)** object representing the resource groups in the active project. Read-only **List**.

## Syntax

 _expression_. **ResourceGroupList**

 _expression_A variable that represents a  **Project** object.


## Example

The following example lists all the resource filters in the active project.


```
Sub SeeAllResGroups() 
 
 Dim Temp As Variant 
 Dim ResGroupNames As String 
 
 For Each Temp In ActiveProject.ResourceGroupList 
 ResGroupNames = ResGroupNames &amp; vbCrLf &amp; Temp 
 Next Temp 
 
 MsgBox ResGroupNames 
 
End Sub
```

