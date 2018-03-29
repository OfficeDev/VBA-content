---
title:  Application.FileOpenEx Method (Project)
keywords: vbapj.chm102
f1_keywords:
- vbapj.chm102
ms.prod: project-server
api_name:
- Project.Application.FileOpenEx
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
Method name. For example,  Application.AddNewColumn Method (Project)
-->

# method name

<!--
{One sentence description of method}
-->

## Syntax
<!--
First expression is the syntax of method call
Second expression is the object that exposes the method being called
-->
 _expression_. **<method>**( ** _<argument name>_**, ** _<argument name>_**, ... )

 _expression_ A variable that represents an **<object name>** object.


### Parameters

<!--
List the parameters of the method in a table. 

For example: 

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|-----|-----|-----|-----|
| _Name_|Optional|**String**|The name of the project file, source file, or data source to open. If  _Name_ is not specified, Project displays the **Open** dialog box.|
-->


### Return Value

 **<VBA type>**


## Remarks

<!-- 
{Describe the behavior of the method. Be sure to describe best practices and non-intuitive behavior


For example: 
Using the  **FileOpenEx** method without specifying any arguments displays the **Open** dialog box with the list of enterprise projects if Project Professional is logged on Project Server. Using `FileOpenEx DoNotLoadFromEnterprise:=True` displays the **Open** dialog box for project files on the local computer.

If you use the  **FileOpenEx** method to open a project that is published to Project Server, it opens the file from the Draft database. For example, to programmatically open a project named Project1 as read/write from Project Server, use the following command: `Application.FileOpenEx Name:="<>\Project1"`.

If you do not want to modify a project, set the  _ReadOnly_ parameter to **True**. For example, to open Project2 as read-only, use the following command: `Application.FileOpenEx Name:="<>\Project2", ReadOnly:=True`. To save the file in the Draft database, use the  **Application.FileSave** method. To publish the file from the Draft to the Published database, so that changes are shown to other users, use the **Application.Publish** method.

The  _Name_ parameter can contain a file name string or an ODBC data source name (DSN) and project name string. The syntax for a data source is <DataSourceName>\Projectname. The less than (<) and greater than (>) symbols must be included, and a backslash ( \ ) must separate the data source name from the project name. _DataSourceName_ itself can either be one of the ODBC data source names installed on the computer or a path and file name for a file-based database.

-->
