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
<!--
The example YAML block above this comment: 
title: <methodname method (workload)>
keywords: <assigned by VBA product team.>
f1_keywords: <assigned by VBA product team>
ms.prod: name of product that hosts this VBA code
ms.date: The date that the topic is checked in to master branch for publication 
-->


<!-- 
Property name. For example,  AccessObject.Name Property (Access)
-->

# property name (client)

<!-- 
Description of property return value
For example: 
Returns a  **Date** indicating the date and time when the design of the specified object was last modified. Read-only.

-->
Returns a {describe what property returns}


## Syntax
<!--
First expression is the syntax of property read
Second expression is the object that exposes the property being read
-->

 _expression_. **{property name}**

 _expression_ A variable that represents an **{exposing VBA object}** object.


## Example

<!-- 
Show an example of the property being read. Include description of example

For example: 
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
-->

## See also
<!-- 
Optional:  Link to relevant API or conceptual articles
-->

#### Concepts

<!-- 
Link to the parent VBA object

For example: 
[AccessObject Object](accessobject-object-access.md)
-->
