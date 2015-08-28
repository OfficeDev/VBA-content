
# Application.OptionsInterfaceEx Method (Project)

 **Last modified:** July 28, 2015

Sets some display options and Project Guide options.

## Syntax

 _expression_. **OptionsInterfaceEx**( **_ShowResourceAssignmentIndicators_**,  **_ShowEditToStartFinishDates_**,  **_ShowEditsToWorkUnitsDurationIndicators_**,  **_ShowDeletionInNameColumn_**,  **_DisplayProjectGuide_**,  **_ProjectGuideUseDefaultFunctionalLayoutPage_**,  **_ProjectGuideFunctionalLayoutPage_**,  **_ProjectGuideUseDefaultContent_**,  **_ProjectGuideContent_**,  **_SetAsDefaults_**,  **_UseOMIDs_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShowResourceAssignmentIndicators|Optional| **Boolean**| **True** if Project displays indicators and options buttons for resource assignments. The default value is **False**.|
|ShowEditToStartFinishDates|Optional| **Boolean**| **True** if Project displays actions on the undo stack for edits to start and finish dates. The default value is **False**.|
|ShowEditsToWorkUnitsDurationIndicators|Optional| **Boolean**| **True** if Project displays actions on the undo stack for edits to duration, work, or units. The default value is **False**.|
|ShowDeletionInNameColumn|Optional| **Boolean**| **True** if Project displays actions on the undo stack upon deletion of a value in the **Task Name** or **Resource Name** field. The default value is **False**.|
|DisplayProjectGuide|Optional| **Boolean**| **True** if the Project Guide should be shown by default on startup and for all new projects. The default value is **False**.|
|ProjectGuideUseDefaultFunctionalLayoutPage|Optional| **Boolean**| **True** if the Project Guide uses default content. **False** if you want to use custom content for the Project Guide. The default value is **True**.|
|ProjectGuideFunctionalLayoutPage|Optional| **String**|The URL or path and file name for the XML file used for custom content in the  **Project Guide**.|
|ProjectGuideUseDefaultContent|Optional| **Boolean**| **True** if the **Project Guide** uses default content. **False** if you want to use custom content for the Project Guide. The default value is **True**.|
|ProjectGuideContent|Optional| **String**|The URL or path and file name for the XML file used for custom content in the Project Guide.|
|SetAsDefaults|Optional| **Boolean**| **True** if the Project Guide settings for the active project should be used as the default for all new projects. The default value is **False**.|
|UseOMIDs|Optional| **Variant**| **True** if Project uses internal IDs to match different-language or renamed Organizer items between projects. The default is **True**. See also the  ** [UseOMIDs](15339e09-0b65-d939-df47-eb538dee7c38.md)** property.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, the default value is specified by the setting on the  **Display** tab of the **Project Options** dialog box. TheUseOMIDs default value is the **Use internal IDs** option on the **Advanced** tab.


 **Note**  The  **Project Options** dialog box does not include settings for the Project Guide, which is deprecated in Project. Project Guide options can only be set programmatically, for using custom project guides. Instead of creating new project guide content, developers should create task pane apps.

Using the  **OptionsInterfaceEx** method with no arguments displays the **Project Options** dialog box with the **General** tab selected. The **OptionsInterfaceEx** method is not available when a report view is active.

