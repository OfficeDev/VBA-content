
# Project.SetObjectMatchingID Method (Project)

 **Last modified:** July 28, 2015

Sets the matching identification value of an object in the  **Organizer** dialog box, for example to change the view specified by "Gantt Chart".

## Syntax

 _expression_. **SetObjectMatchingID**( **_ObjectType_**,  **_ObjectName_**,  **_MatchingID_**)

 _expression_A variable that represents a  **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ObjectType|Required| **Long**|The type of object, specified by a  ** [pjOrganizer](d176be88-4df9-3826-c806-f7f650fffb39.md)** constant.|
|ObjectName|Required| **String**|Display name of the object.|
|MatchingID|Required| **String**|String specifying the matching ID to set.|

## Example

The following example sets the matching ID of a  **pjView** object type with the display name "Gantt Chart" to "Gantt Chart 1".


```
ActiveProject.SetObjectMatchingID ObjectType:=pjView, ObjectName:="Gantt Chart", MatchingID:="Gantt Chart 1"
```

