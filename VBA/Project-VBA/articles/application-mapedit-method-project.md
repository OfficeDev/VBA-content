---
title: Application.MapEdit Method (Project)
keywords: vbapj.chm243
f1_keywords:
- vbapj.chm243
ms.prod: project-server
api_name:
- Project.Application.MapEdit
ms.assetid: 316d596e-95b3-d616-c8d6-21da651ff284
ms.date: 06/08/2017
---


# Application.MapEdit Method (Project)

Creates or edits an import/export map.


## Syntax

 _expression_. **MapEdit**( ** _Name_**, ** _Create_**, ** _OverwriteExisting_**, ** _NewName_**, ** _DataCategory_**, ** _CategoryEnabled_**, ** _TableName_**, ** _FieldName_**, ** _ExternalFieldName_**, ** _ExportFilter_**, ** _ImportMethod_**, ** _MergeKey_**, ** _HeaderRow_**, ** _AssignmentData_**, ** _TextDelimiter_**, ** _TextFileOrigin_**, ** _UseHtmlTemplate_**, ** _TemplateFile_**, ** _IncludeImage_**, ** _ImageFile_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the map to create, copy, or edit.|
| _Create_|Optional|**Boolean**|**True** if Project should create a new map. If NewName is not specified, the new map is given the name specified with Name. Otherwise, the new map is a copy of the map specified with Name and is given the name specified with NewName. The default value is **False**.|
| _OverwriteExisting_|Optional|**Boolean**|**True** if an existing map should be overwritten with a new one. The default value is **False**.|
| _NewName_|Optional|**String**|A new name for the existing map (Create is  **False** ) or the name for the new map copied from the existing map (Create is **True** ). If NewName is not specified and Create is **False**, the map specified with Name retains its current name. The default value is an empty string ("").|
| _DataCategory_|Optional|**Long**| The category of data that will be modified by other arguments. Required if any of CategoryEnabled, TableName, FieldName, ExternalFieldName, ExportFilter, or MergeKey are specified. Can be one of the following **[PjDataCategories](pjdatacategories-enumeration-project.md)** constants: **pjMapTasks**, **pjMapResources**, or **pjMapAssignments**.|
| _CategoryEnabled_|Optional|**Boolean**|**True** if the map imports and exports the category of data specified with DataCategory. If Create is **True** and NewName is not specified, CategoryEnabled is set to **True**.|
| _TableName_|Optional|**String**|The name of the external table or worksheet that the map imports data from or exports data to. The type of table is determined by the value of DataCategory. If Create is  **True** and NewName is not specified, TableName is required.|
| _FieldName_|Optional|**String**|The name of a field to add to the map. The field is mapped to the external field specified with ExternalFieldName. The type of field is determined by the value of DataCategory. If Create is  **True** and NewName is not specified, FieldName is required.|
| _ExternalFieldName_|Optional|**String**|The name of the external field to add to the map. The external field is mapped to the field specified with FieldName. If ExternalFieldName is not specified, the name specified with FieldName is also used for ExternalFieldName.|
| _ExportFilter_|Optional|**String**|The name of the filter to use when exporting data. The type of filter is determined by the value of DataCategory. The default value is "All Tasks" when DataCategory is  **pjMapTasks**, "All Resources" when DataCategory is **pjMapResources**, and ExportFilter is ignored when DataCategory is **pjMapAssignments**.|
| _ImportMethod_|Optional|**Long**|The method to use when importing data. Can be one of the  **[PjImportMethods](pjimportmethods-enumeration-project.md)** constants. The default value is **pjImportNew**.|
| _MergeKey_|Optional|**String**|The name of the project field to use as a key when merging imported data. The field must exist and have already been added to the map. The type of field is determined by the value of DataCategory. If ImportMethod is  **pjImportMerge**, MergeKey is required.|
| _HeaderRow_|Optional|**Boolean**|**True** if a column header row should be created in the external file during an export and whether it exists in the external file during an import. If creating a headerless map (HeaderRow is **False** ) that will be used to import the same data it exports, ExternalFieldName is required and must be a sequentially numbered value for each field exported, beginning with "1", to indicate its column position in the exported file. The default value is **True**.|
| _AssignmentData_|Optional|**Boolean**|**True** if assignment rows should be included with exported resources and tasks. If **True**, assigned resources appear under each task in a task table and assigned tasks appear under each resource in a resource table. Data exported when AssignmentData is **True** cannot be imported by Project. The default value is **False**.|
| _TextDelimiter_|Optional|**String**|The character to use as a field delimiter when importing data from a text file. The default value is a tab character.|
| _TextFileOrigin_|Optional|**Long**|Specifies the character set under which a text file was created. Can be one of the following  **[PjTextFileOrigin](pjtextfileorigin-enumeration-project.md)** constants: **pjTextOriginWin**, **pjTextOriginDOS**, **pjTextOriginUnicode**, or **pjTextOriginMac**.|
| _UseHtmlTemplate_|Optional|**Boolean**|**True** if an export to an HTML file will be based on an HTML template.|
| _TemplateFile_|Optional|**String**|The HTML template file to use when exporting to HTML. If UseHtmlTemplate is  **True** and the map specified with Name does not contain the name of an HTML template file, TemplateFile is required.|
| _IncludeImage_|Optional|**Boolean**|**True** if a reference to an image file should be included when exporting to HTML. The default value is **False**.|
| _ImageFile_|Optional|**String**|The name of an image file to include when exporting to HTML.|

### Return Value

 **Boolean**


## Example

The following example creates a simple map that allows the information on the default Gantt Chart to be exported and imported.


```vb
Sub MakeEntryTableMap() 
 
 MapEdit Name:="Fields in the Gantt Chart View", Create:=True, OverwriteExisting:=True, _ 
 DataCategory:=pjMapTasks, CategoryEnabled:=True, TableName:="Task_Table", _ 
 FieldName:="ID", ExternalFieldName:="ID" 
 MapEdit Name:="Fields in the Gantt Chart View", DataCategory:=pjMapTasks, _ 
 FieldName:="Name", ExternalFieldName:="Tasks" 
 MapEdit Name:="Fields in the Gantt Chart View", DataCategory:=pjMapTasks, _ 
 FieldName:="Duration" 
 MapEdit Name:="Fields in the Gantt Chart View", DataCategory:=pjMapTasks, _ 
 FieldName:="Start", ExternalFieldName:="Start_Date" 
 MapEdit Name:="Fields in the Gantt Chart View", DataCategory:=pjMapTasks, _ 
 FieldName:="Finish", ExternalFieldName:="Finish_Date" 
 MapEdit Name:="Fields in the Gantt Chart View", DataCategory:=pjMapTasks, _ 
 FieldName:="Predecessors" 
 MapEdit Name:="Fields in the Gantt Chart View", DataCategory:=pjMapTasks, _ 
 FieldName:="Resource Names", ExternalFieldName:="Resources" 
 
End Sub
```


