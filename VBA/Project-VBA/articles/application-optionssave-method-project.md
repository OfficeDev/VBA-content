---
title: Application.OptionsSave Method (Project)
keywords: vbapj.chm650
f1_keywords:
- vbapj.chm650
ms.prod: project-server
api_name:
- Project.Application.OptionsSave
ms.assetid: 658a4b31-8bd6-8dbb-852f-a7f604386215
ms.date: 06/08/2017
---


# Application.OptionsSave Method (Project)

Sets save options for project files.


## Syntax

 _expression_. **OptionsSave**( ** _DefaultSaveFormat_**, ** _DefaultProjectsPath_**, ** _DefaultUserTemplatesPath_**, ** _DefaultWorkgroupTemplatesPath_**, ** _ExpandDatabaseTimephasedData_**, ** _AutomaticSave_**, ** _AutomaticSaveInterval_**, ** _AutomaticSaveOptions_**, ** _AutomaticSavePrompt_**, ** _SetDefaultsDatabase_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultSaveFormat_|Optional|**String**|Specifies the default format when saving a file. Can be one of the following strings: "MSProject.mpp", "MSProject.mpt", "MSProject.mpp12", or "MSProject.mpp9".|
| _DefaultProjectsPath_|Optional|**String**|Specifies the default location for project files.|
| _DefaultUserTemplatesPath_|Optional|**String**|Specifies the default location for user templates.|
| _DefaultWorkgroupTemplatesPath_|Optional|**String**|Specifies the default location for workgroup templates.|
| _ExpandDatabaseTimephasedData_|Optional|**Boolean**|**True** if timephased data should be expanded to a readable format when saving to a database. **False** if timephased data should remain in a compressed binary format. The default value is **False**.|
| _AutomaticSave_|Optional|**Boolean**|**True** if Project automatically saves files.|
| _AutomaticSaveInterval_|Optional|**Long**|Specifies how often (in minutes) Project automatically saves.|
| _AutomaticSaveOptions_|Optional|**Long**|Specifies whether Project saves only the active file or all changed files. Can be one of the following  **[PjAutomaticSaveOptions](pjautomaticsaveoptions-enumeration-project.md)** constants.|
| _AutomaticSavePrompt_|Optional|**Boolean**|**True** if alerts display when automatically saving files.|
| _SetDefaultsDatabase_|Optional|**Boolean**|**True** if the value specified in the **Database save options** section, found on the **Save** tab of the **Options** dialog box, is used as the default value for new projects. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, the default value is specified by the corresponding setting on the  **Save** tab of the **Project Options** dialog box.

Using the  **OptionsSave** method without specifying any arguments displays the **Project Options** dialog box with the **General** tab selected.


## Example

The following example turns off the automatic saving feature.


```vb
Sub Options_Save() 
    OptionsSave AutomaticSave:=False 
End Sub
```


