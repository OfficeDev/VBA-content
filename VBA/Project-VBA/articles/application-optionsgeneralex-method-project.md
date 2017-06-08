---
title: Application.OptionsGeneralEx Method (Project)
keywords: vbapj.chm642
f1_keywords:
- vbapj.chm642
ms.prod: project-server
api_name:
- Project.Application.OptionsGeneralEx
ms.assetid: c82b09d5-0937-ed06-58d6-e6b5fda186ac
ms.date: 06/08/2017
---


# Application.OptionsGeneralEx Method (Project)

Sets some options that are on the  **General**,  **Schedule**, and  **Advanced** tabs of the **Project Options** dialog box.


## Syntax

 _expression_. **OptionsGeneralEx**( ** _PlanningWizard_**, ** _WizardUsage_**, ** _WizardErrors_**, ** _WizardScheduling_**, ** _ShowTipOfDay_**, ** _AutoAddResources_**, ** _StandardRate_**, ** _OvertimeRate_**, ** _LastFile_**, ** _SummaryInfo_**, ** _UserName_**, ** _SetDefaults_**, ** _ShowWelcome_**, ** _AutoFilter_**, ** _MacroVirusProtection_**, ** _DisplayRecentFiles_**, ** _RecentFilesMaximum_**, ** _FontConversion_**, ** _ShowStartupWorkpane_**, ** _MaxUndoRecords_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PlanningWizard_|Optional|**Boolean**|**True** if the Planning Wizard is active. Planning Wizard settings are on the **Advanced** tab ofn the **Project Options** dialog box.|
| _WizardUsage_|Optional|**Boolean**|**True** if the Planning Wizard displays tips about using Project more effectively.|
| _WizardErrors_|Optional|**Boolean**|**True** if the Planning Wizard displays messages about errors.|
| _WizardScheduling_|Optional|**Boolean**|**True** if the Planning Wizard displays messages about scheduling problems.|
| _ShowTipOfDay_|Optional||Because of changes in the Project object model, this argument no longer has an effect. It is retained for backward compatibility.|
| _AutoAddResources_|Optional|**Boolean**|**True** if resources are automatically added to the resource pool.|
| _StandardRate_|Optional|**Variant**|The default standard pay rate for resources.|
| _OvertimeRate_|Optional|**Variant**|The default overtime pay rate for resources.|
| _LastFile_|Optional|**Boolean**|**True** if the last opened file is automatically opened when Project starts.|
| _SummaryInfo_|Optional|**Boolean**|**True** if the **Project Information** dialog box appears when a new project is created.|
| _UserName_|Optional|**String**|The name of the current user.|
| _SetDefaults_|Optional|**Boolean**|**True** if the values of AutoAddResources, StandardRate, and OvertimeRate are used as default values for new projects.|
| _ShowWelcome_|Optional||Because of changes in the Project object model, this argument no longer has an effect. It is retained for backward compatibility.|
| _AutoFilter_|Optional|**Boolean**|**True** if the AutoFilter is active.|
| _MacroVirusProtection_|Optional||Because of changes in the Project object model, this argument no longer has an effect. It is retained for backward compatibility.|
| _DisplayRecentFiles_|Optional|**Boolean**|**True** if a list of recently used files appears on the **File** menu.|
| _RecentFilesMaximum_|Optional|**Integer**|The maximum number of recently used files to display on the  **File** menu. Can be a number from 0 to 9. Setting RecentFilesMaximum to 0 also sets DisplayRecentFiles to **False**.|
| _FontConversion_|Optional|**Boolean**|**True** if the font automatically changes when opening a file that uses a font that cannot display native characters. The FontConversion argument is ignored unless an East Asian version of Project is used.|
| _ShowStartupWorkpane_|Optional||Because of changes in the Project object model, this argument no longer has an effect. It is retained for backward compatibility.|
| _MaxUndoRecords_|Optional|**Variant**|The maximum number of records stored in the undo stack.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, the default value is specified by the corresponding settings in the  **Project Options** dialog box.

Using the  **OptionsGeneralEx** method without specifying any arguments displays the **Project Options** dialog box with the **General** tab selected.


