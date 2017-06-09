---
title: Application.OptionsInterfaceEx Method (Project)
keywords: vbapj.chm651
f1_keywords:
- vbapj.chm651
ms.prod: project-server
api_name:
- Project.Application.OptionsInterfaceEx
ms.assetid: da4dc69c-021f-7ecb-22f6-aebf1d9252dd
ms.date: 06/08/2017
---


# Application.OptionsInterfaceEx Method (Project)

Sets some display options and Project Guide options.


## Syntax

 _expression_. **OptionsInterfaceEx**( ** _ShowResourceAssignmentIndicators_**, ** _ShowEditToStartFinishDates_**, ** _ShowEditsToWorkUnitsDurationIndicators_**, ** _ShowDeletionInNameColumn_**, ** _DisplayProjectGuide_**, ** _ProjectGuideUseDefaultFunctionalLayoutPage_**, ** _ProjectGuideFunctionalLayoutPage_**, ** _ProjectGuideUseDefaultContent_**, ** _ProjectGuideContent_**, ** _SetAsDefaults_**, ** _UseOMIDs_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShowResourceAssignmentIndicators_|Optional|**Boolean**|**True** if Project displays indicators and options buttons for resource assignments. The default value is **False**.|
| _ShowEditToStartFinishDates_|Optional|**Boolean**|**True** if Project displays actions on the undo stack for edits to start and finish dates. The default value is **False**.|
| _ShowEditsToWorkUnitsDurationIndicators_|Optional|**Boolean**|**True** if Project displays actions on the undo stack for edits to duration, work, or units. The default value is **False**.|
| _ShowDeletionInNameColumn_|Optional|**Boolean**|**True** if Project displays actions on the undo stack upon deletion of a value in the **Task Name** or **Resource Name** field. The default value is **False**.|
| _DisplayProjectGuide_|Optional|**Boolean**|**True** if the Project Guide should be shown by default on startup and for all new projects. The default value is **False**.|
| _ProjectGuideUseDefaultFunctionalLayoutPage_|Optional|**Boolean**|**True** if the Project Guide uses default content. **False** if you want to use custom content for the Project Guide. The default value is **True**.|
| _ProjectGuideFunctionalLayoutPage_|Optional|**String**|The URL or path and file name for the XML file used for custom content in the  **Project Guide**.|
| _ProjectGuideUseDefaultContent_|Optional|**Boolean**|**True** if the **Project Guide** uses default content. **False** if you want to use custom content for the Project Guide. The default value is **True**.|
| _ProjectGuideContent_|Optional|**String**|The URL or path and file name for the XML file used for custom content in the Project Guide.|
| _SetAsDefaults_|Optional|**Boolean**|**True** if the Project Guide settings for the active project should be used as the default for all new projects. The default value is **False**.|
| _UseOMIDs_|Optional|**Variant**|**True** if Project uses internal IDs to match different-language or renamed Organizer items between projects. The default is **True**. See also the **[UseOMIDs](application-useomids-property-project.md)** property.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, the default value is specified by the setting on the  **Display** tab of the **Project Options** dialog box. The _UseOMIDs_ default value is the **Use internal IDs** option on the **Advanced** tab.


 **Note**  The  **Project Options** dialog box does not include settings for the Project Guide, which is deprecated in Project. Project Guide options can only be set programmatically, for using custom project guides. Instead of creating new project guide content, developers should create task pane apps.

Using the  **OptionsInterfaceEx** method with no arguments displays the **Project Options** dialog box with the **General** tab selected. The **OptionsInterfaceEx** method is not available when a report view is active.


