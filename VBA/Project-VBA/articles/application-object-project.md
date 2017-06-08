---
title: Application Object (Project)
ms.prod: project-server
api_name:
- Project.Application
ms.assetid: 8eb91712-7784-a102-38c0-19bb056c27e9
ms.date: 06/08/2017
---


# Application Object (Project)

Represents the entire Project application. The  **Application** object contains:


- Application-wide settings and options (many of the options in the  **Options** dialog box on the **Tools** menu, for example).
    
- Properties that return top-level objects, such as  **ActiveCell**, **ActiveProject**, and so forth.
    
- Methods that act on application-wide elements, such as views, selections, editing actions, and so forth.
    

## Using the Application Object

Use the  **[Application](http://msdn.microsoft.com/library/935ad507-7df9-ce7b-16ab-4270349d9b74%28Office.15%29.aspx)** property to return an **Application** object in Project . The following example applies the **Windows** property to the **Application** object.


```
Application.Windows("Project1.mpp").Activate
```


## Using Project From Another Application: Late Binding

The following example creates the Microsoft Project  **Application** object at run time, creates a new project, adds a task, saves the project, and then closes the Project . For example, copy and paste the **CreateProject_Late** macro to the **ThisDocument** module in the Visual Basic Editor (VBE) of Word.


 **Note**  Because the application queries the  **MSProject.Application** type library only at run time, Microsoft IntelliSense is not available and performance is relatively poor with late binding. Scripting languages, such as JavaScript and VBScript, require late binding. VBScript supports only the generic **Object** and **Variant** data types. For better performance in VBA and other compiled languages, you should use early binding by setting a reference to the Project type library.


```
Sub CreateProject_Late() 
    Dim pjApp As Object 
    Set pjApp = CreateObject("MSProject.Application") 
    pjApp.Visible = True 
    pjApp.FileNew 
    pjApp.ActiveProject.Tasks.Add "Hang clocks" 
    pjApp.FileSaveAs "Clocks.mpp" 
    pjApp.FileClose 
    pjApp.Quit 
End Sub
```

If you do not set the  **Visible** property to **True**, the Project application operates in the background without being visible.


## Using Project From Another Application: Early Binding

Early binding has better performance because it loads the type library at design time. To use early binding, you must set a reference to the Project application from the application you are working in. For example, in the VBE for a Word document, click  **References** on the **Tools** menu, scroll through the **Available References** list, and then choose the **Microsoft Project 15.0 Object Library** checkbox.

The following example opens a project from another application such as Excel , adds a task, and then saves and closes the project. 




```
Sub ModifyProject_Early() 
    Dim pjApp As MSProject.Application 
    Set pjApp = New MSProject.Application 
    pjApp.Visible = True 
    pjApp.FileOpen "Clocks.mpp" 
    pjApp.ActiveProject.Tasks.Add "Wind clocks" 
    pjApp.FileSave 
    pjApp.FileClose 
    pjApp.Quit 
End Sub
```


## Remarks




 **Important**  For application-level events, register event handlers  _after_ you set `Application.Visible = True`.



If you instantiate Project from another application and register an application-level event before setting the  **Visible** property of the **Application** object to **True**, the properties and methods of child objects of **Application** do not work. For example, `Application.ActiveProject.Name` is not accessible.

Many of the properties and methods that return the most common user-interface objects, such as the active project—represented by the  **[ActiveProject](http://msdn.microsoft.com/library/07844166-ca9b-15eb-a5e2-6f00a7c0a030%28Office.15%29.aspx)** property—can be used without the **Application** object qualifier. For example, instead of writing `Application.ActiveProject.Visible = True` you can write `ActiveProject.Visible = True`


## Events



|**Name**|
|:-----|
|[AfterCubeBuilt](http://msdn.microsoft.com/library/f57a3391-dbbe-42eb-cf99-205b754c7cc1%28Office.15%29.aspx)|
|[ApplicationBeforeClose](http://msdn.microsoft.com/library/9523a793-b4c1-fd79-303e-b167d7f80025%28Office.15%29.aspx)|
|[ConnectionStatusChanged](http://msdn.microsoft.com/library/ffc6fc8a-f5b7-3a3d-4829-712a8305ed17%28Office.15%29.aspx)|
|[IsFunctionalitySupported](http://msdn.microsoft.com/library/f6462a3b-5a36-3b2e-79bd-78cce567aed8%28Office.15%29.aspx)|
|[JobCompleted](http://msdn.microsoft.com/library/44f7987c-92e0-a302-a775-7e62dab2ef86%28Office.15%29.aspx)|
|[JobStart](http://msdn.microsoft.com/library/874b35cb-bb90-b8dc-3c22-84c8809c3177%28Office.15%29.aspx)|
|[LoadWebPage](http://msdn.microsoft.com/library/393115c4-6245-3a1a-3c98-a5ddc1416aa0%28Office.15%29.aspx)|
|[LoadWebPane](http://msdn.microsoft.com/library/b9fefabb-3d0b-9aa7-6d3b-b8fd8000571d%28Office.15%29.aspx)|
|[NewProject](http://msdn.microsoft.com/library/de3c9e06-405a-8f63-6210-013f5d292c20%28Office.15%29.aspx)|
|[OnUndoOrRedo](http://msdn.microsoft.com/library/7f60e893-81d0-1b2f-c5f5-ec1451633fa7%28Office.15%29.aspx)|
|[PaneActivate](http://msdn.microsoft.com/library/8230c818-6df3-bbdc-5e71-0e6e6b03e172%28Office.15%29.aspx)|
|[ProjectAfterSave](http://msdn.microsoft.com/library/e0dbe6de-0b5e-1b4a-2b30-8c228249b491%28Office.15%29.aspx)|
|[ProjectAssignmentNew](http://msdn.microsoft.com/library/dcb4acc6-a113-1e93-5f08-e9e68b902b96%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentChange](http://msdn.microsoft.com/library/9d94303c-f8f6-1681-0829-23f240afc570%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentChange2](http://msdn.microsoft.com/library/99fce7af-00de-42d8-4b61-e97774cc19ed%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentDelete](http://msdn.microsoft.com/library/f0db513e-3dec-e9d6-8385-ac0117e8f28e%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentDelete2](http://msdn.microsoft.com/library/2753a140-e01b-b2c1-233f-f9f265737b47%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentNew](http://msdn.microsoft.com/library/5caedd9a-94b1-daa6-762a-a037dae4f917%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentNew2](http://msdn.microsoft.com/library/9e2f3358-325e-53b9-3da6-5323482e2a47%28Office.15%29.aspx)|
|[ProjectBeforeClearBaseline](http://msdn.microsoft.com/library/4aa11658-7962-a46f-c914-5ed3bebd15a3%28Office.15%29.aspx)|
|[ProjectBeforeClose](http://msdn.microsoft.com/library/90e75c72-03f9-25ab-1339-94d9ff8933a2%28Office.15%29.aspx)|
|[ProjectBeforeClose2](http://msdn.microsoft.com/library/24b43d85-f99c-915c-47fe-0df5875fc479%28Office.15%29.aspx)|
|[ProjectBeforePrint](http://msdn.microsoft.com/library/7cc8de23-c3e3-81df-ae26-37c4e639dd81%28Office.15%29.aspx)|
|[ProjectBeforePrint2](http://msdn.microsoft.com/library/93e243b7-d765-e3d9-d061-dd98407010d1%28Office.15%29.aspx)|
|[ProjectBeforePublish](http://msdn.microsoft.com/library/5778ec6c-a8c0-0a05-145c-c9ad6132bf87%28Office.15%29.aspx)|
|[ProjectBeforeResourceChange](http://msdn.microsoft.com/library/d676f2c7-8857-70d7-41c6-4c505a0bcbcc%28Office.15%29.aspx)|
|[ProjectBeforeResourceChange2](http://msdn.microsoft.com/library/84128c94-0d0d-f8f2-6d5a-4c05a61a0a8d%28Office.15%29.aspx)|
|[ProjectBeforeResourceDelete](http://msdn.microsoft.com/library/aadef12e-57dc-210e-d29a-54f79d1c1abd%28Office.15%29.aspx)|
|[ProjectBeforeResourceDelete2](http://msdn.microsoft.com/library/3665f6e0-6df8-0a8d-28c1-49bfe51ffad5%28Office.15%29.aspx)|
|[ProjectBeforeResourceNew](http://msdn.microsoft.com/library/a432c713-d1fa-0743-ff4e-90fbd724dfe4%28Office.15%29.aspx)|
|[ProjectBeforeResourceNew2](http://msdn.microsoft.com/library/24c28eac-946b-80fb-5dcb-8b9ef499b547%28Office.15%29.aspx)|
|[ProjectBeforeSave](http://msdn.microsoft.com/library/406986e7-22f6-109e-1973-f22e81081111%28Office.15%29.aspx)|
|[ProjectBeforeSave2](http://msdn.microsoft.com/library/5afcdb4c-85e6-183c-f6e7-333d2a7ea3d4%28Office.15%29.aspx)|
|[ProjectBeforeSaveBaseline](http://msdn.microsoft.com/library/bcdd2134-03dd-e26d-66db-095bda6a7162%28Office.15%29.aspx)|
|[ProjectBeforeTaskChange](http://msdn.microsoft.com/library/995024c3-b031-0ddd-0fbe-4d817f237473%28Office.15%29.aspx)|
|[ProjectBeforeTaskChange2](http://msdn.microsoft.com/library/00992e39-dcbd-3826-4ce6-e2be55dc9c2c%28Office.15%29.aspx)|
|[ProjectBeforeTaskDelete](http://msdn.microsoft.com/library/3acc4ba4-0fdc-61fd-17df-e6450055a39b%28Office.15%29.aspx)|
|[ProjectBeforeTaskDelete2](http://msdn.microsoft.com/library/2c695579-bfe4-d109-eebc-4fb258a95c1e%28Office.15%29.aspx)|
|[ProjectBeforeTaskNew](http://msdn.microsoft.com/library/77418f84-1d82-b227-75f8-c688b7bddf82%28Office.15%29.aspx)|
|[ProjectBeforeTaskNew2](http://msdn.microsoft.com/library/4df0eb83-e60d-943d-aecf-57a2f857ae42%28Office.15%29.aspx)|
|[ProjectCalculate](http://msdn.microsoft.com/library/44dbf3f9-4a7d-2e85-aa63-915ea47af008%28Office.15%29.aspx)|
|[ProjectResourceNew](http://msdn.microsoft.com/library/9b030fbc-5cca-df10-f7a3-613d7ad70dc7%28Office.15%29.aspx)|
|[ProjectTaskNew](http://msdn.microsoft.com/library/40e9d8da-f863-a73e-56e9-bb89327142fb%28Office.15%29.aspx)|
|[SaveCompletedToServer](http://msdn.microsoft.com/library/05ca27a0-a6cd-efbd-eff8-4f457c3de5c0%28Office.15%29.aspx)|
|[SaveStartingToServer](http://msdn.microsoft.com/library/e9d19b19-b916-a85d-486a-4a8676998b6c%28Office.15%29.aspx)|
|[SecondaryViewChange](http://msdn.microsoft.com/library/f0f3f81b-c75f-79ee-db8b-6bdd32a3702f%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/b54d0956-7eab-db5f-394a-5120bc111afd%28Office.15%29.aspx)|
|[WindowBeforeViewChange](http://msdn.microsoft.com/library/c3eb450d-2a74-6ae1-175c-1d61c90b22ca%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/141940d7-f117-d3a8-2aa5-83679a5fbfd4%28Office.15%29.aspx)|
|[WindowGoalAreaChange](http://msdn.microsoft.com/library/1ae33d11-f8aa-e1a2-b59d-9736ce4a6283%28Office.15%29.aspx)|
|[WindowSelectionChange](http://msdn.microsoft.com/library/239c0a87-7966-b4b5-5731-9fe059f56a43%28Office.15%29.aspx)|
|[WindowSidepaneDisplayChange](http://msdn.microsoft.com/library/8c4c22f4-4005-eff5-2964-880982634e78%28Office.15%29.aspx)|
|[WindowSidepaneTaskChange](http://msdn.microsoft.com/library/674a8134-1e34-2658-6c67-5eb92c628ed8%28Office.15%29.aspx)|
|[WindowViewChange](http://msdn.microsoft.com/library/e6a5f884-5bb9-f975-9237-25996b436589%28Office.15%29.aspx)|
|[WorkpaneDisplayChange](http://msdn.microsoft.com/library/8fad51ed-57f5-a34d-6ef6-f699b605c10c%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[About](http://msdn.microsoft.com/library/323c2400-e886-300a-f8ad-a4fed3fe00bf%28Office.15%29.aspx)|
|[ActivateMicrosoftApp](http://msdn.microsoft.com/library/a9b59db3-7ad2-8674-9026-090e161ef983%28Office.15%29.aspx)|
|[AddNewColumn](http://msdn.microsoft.com/library/009071ad-b713-4252-ab1c-781d58620d8c%28Office.15%29.aspx)|
|[AddProgressLine](http://msdn.microsoft.com/library/f7a780f6-63af-e495-9fce-f3f1031bdfa0%28Office.15%29.aspx)|
|[AddResourcesFromProjectServer](http://msdn.microsoft.com/library/74fe4224-0019-5daa-11ae-3bdd6f2f5abb%28Office.15%29.aspx)|
|[AddSiteColumn](http://msdn.microsoft.com/library/0ec78b0b-b4bf-3dea-0ed6-af78798bd7cd%28Office.15%29.aspx)|
|[AfterUnloadWebBrowserControl](http://msdn.microsoft.com/library/794718d0-2f23-06ad-1d14-19fb7e946a1f%28Office.15%29.aspx)|
|[Alerts](http://msdn.microsoft.com/library/58c935d9-35a3-953b-4003-dc88f8532854%28Office.15%29.aspx)|
|[AlignTableCellBottom](http://msdn.microsoft.com/library/3eedfcb4-eb75-163f-6c3a-4dde97ddb110%28Office.15%29.aspx)|
|[AlignTableCellTop](http://msdn.microsoft.com/library/51eca157-64c4-f114-243e-895d97adf45a%28Office.15%29.aspx)|
|[AlignTableCellVerticalCenter](http://msdn.microsoft.com/library/c790d8f7-e792-0718-3166-312640ff3f73%28Office.15%29.aspx)|
|[AppExecute](http://msdn.microsoft.com/library/af263a18-9b88-e6c2-d44c-a2ac41951624%28Office.15%29.aspx)|
|[ApplyReport](http://msdn.microsoft.com/library/869640a0-e45e-2e89-e3c9-ca15113ba8d3%28Office.15%29.aspx)|
|[ApplyReportLayoutTemplate](http://msdn.microsoft.com/library/cbc233c9-b955-3cd2-b1b8-99e4257bfea0%28Office.15%29.aspx)|
|[AppMaximize](http://msdn.microsoft.com/library/c194beb5-3d8c-93ac-9338-54d52f6e460a%28Office.15%29.aspx)|
|[AppMinimize](http://msdn.microsoft.com/library/3794f51b-783e-0efa-7bdc-333f2964cf1f%28Office.15%29.aspx)|
|[AppMove](http://msdn.microsoft.com/library/73ab96b7-4985-b25f-d202-89e6230e6e4e%28Office.15%29.aspx)|
|[AppRestore](http://msdn.microsoft.com/library/f50a1158-83d1-e38e-65e6-cdc456f14bc7%28Office.15%29.aspx)|
|[AppSize](http://msdn.microsoft.com/library/31183106-d66d-235d-608c-02d3844c0e1b%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/5d509f1c-2dba-0cd1-540f-3a6aa2a9c1c4%28Office.15%29.aspx)|
|[AutoFilter](http://msdn.microsoft.com/library/391d5a61-cba3-9e28-c448-d0befcc456c7%28Office.15%29.aspx)|
|[AutoSaveToGlobal](http://msdn.microsoft.com/library/8b8d0169-a1c1-8771-bc90-503a17e00b26%28Office.15%29.aspx)|
|[BarBoxFormat](http://msdn.microsoft.com/library/4c491952-533a-21a9-49fc-ccb7a3342370%28Office.15%29.aspx)|
|[BarBoxStyles](http://msdn.microsoft.com/library/a548985d-f5f3-7646-3b05-b00a9232e370%28Office.15%29.aspx)|
|[BarRounding](http://msdn.microsoft.com/library/6f776070-0a37-a72b-8cf8-ea3fd2c3fd06%28Office.15%29.aspx)|
|[BaseCalendarCreate](http://msdn.microsoft.com/library/c9c92dff-255a-041b-c18d-49d6d75884e3%28Office.15%29.aspx)|
|[BaseCalendarDelete](http://msdn.microsoft.com/library/f9583bd7-6ddb-7115-b7ca-c0e4e8b033e1%28Office.15%29.aspx)|
|[BaseCalendarEditDays](http://msdn.microsoft.com/library/3a65015e-c174-985a-5235-099db363c003%28Office.15%29.aspx)|
|[BaseCalendarRename](http://msdn.microsoft.com/library/e895c89f-1a29-0982-a88b-5af662215573%28Office.15%29.aspx)|
|[BaseCalendarReset](http://msdn.microsoft.com/library/43c842b2-146b-f080-f88b-c1e0ef5526d8%28Office.15%29.aspx)|
|[BaseCalendars](http://msdn.microsoft.com/library/5ae675d2-1be3-eb98-6c35-ff36c3fccf30%28Office.15%29.aspx)|
|[BaselineClear](http://msdn.microsoft.com/library/a319fc88-2421-eafa-e498-4a0a5f173394%28Office.15%29.aspx)|
|[BaselineSave](http://msdn.microsoft.com/library/b64967fe-f029-fc32-762a-f81cac405447%28Office.15%29.aspx)|
|[BoxAlign](http://msdn.microsoft.com/library/2b27c9a0-36fa-1bbd-96e3-267b95ad5407%28Office.15%29.aspx)|
|[BoxCellEdit](http://msdn.microsoft.com/library/27063852-3dc4-57b2-c82a-6210674810ca%28Office.15%29.aspx)|
|[BoxCellEditEx](http://msdn.microsoft.com/library/86405780-ea5f-d32b-b2e5-3d3999c1877d%28Office.15%29.aspx)|
|[BoxCellLayout](http://msdn.microsoft.com/library/9b1ab0f5-d3ef-3258-aa01-ae1dea264ec5%28Office.15%29.aspx)|
|[BoxDataTemplate](http://msdn.microsoft.com/library/ce3530d5-6218-b0db-a890-9a80bca5e3db%28Office.15%29.aspx)|
|[BoxFormat](http://msdn.microsoft.com/library/bc2c0b19-c030-3063-4842-cf1bb146f73f%28Office.15%29.aspx)|
|[BoxFormatEx](http://msdn.microsoft.com/library/2cec4b32-3170-8d0b-f73e-5dc64e5ffa68%28Office.15%29.aspx)|
|[BoxGetXPosition](http://msdn.microsoft.com/library/df7a41c8-01df-bd60-0ae1-0fb60cbc3347%28Office.15%29.aspx)|
|[BoxGetYPosition](http://msdn.microsoft.com/library/8284181f-b677-8cc4-8311-23d50987239c%28Office.15%29.aspx)|
|[BoxLayout](http://msdn.microsoft.com/library/4f26f5d1-41f2-56dc-e376-bcedd29613f9%28Office.15%29.aspx)|
|[BoxLayoutEx](http://msdn.microsoft.com/library/40c80e1c-6763-172d-c48a-0ec7c1fa2412%28Office.15%29.aspx)|
|[BoxLinkLabelsShow](http://msdn.microsoft.com/library/8dbb1406-10e8-d096-540a-4c7cfd61a413%28Office.15%29.aspx)|
|[BoxLinks](http://msdn.microsoft.com/library/da12c972-9647-9e1f-2909-1e0a18aff32b%28Office.15%29.aspx)|
|[BoxLinksEx](http://msdn.microsoft.com/library/f6292e01-3f4a-3b83-e86c-2316c83b2509%28Office.15%29.aspx)|
|[BoxLinkStyleToggle](http://msdn.microsoft.com/library/8367a55b-9a7e-3272-49b2-486c0a284f7d%28Office.15%29.aspx)|
|[BoxProgressMarksShow](http://msdn.microsoft.com/library/fd0ff0bd-7069-5e41-fa50-a47a4b09e9f6%28Office.15%29.aspx)|
|[BoxSet](http://msdn.microsoft.com/library/06bcae73-5208-824d-4f55-119f35b37718%28Office.15%29.aspx)|
|[BoxShowHideFields](http://msdn.microsoft.com/library/b100c012-8ab9-2e39-c8c8-569b1498c5da%28Office.15%29.aspx)|
|[BoxStylesEdit](http://msdn.microsoft.com/library/21a15566-3ee2-521a-f813-0f0baa806bfd%28Office.15%29.aspx)|
|[BoxStylesEditEx](http://msdn.microsoft.com/library/8a473e08-7893-6871-d015-23e1791e67e3%28Office.15%29.aspx)|
|[BoxZoom](http://msdn.microsoft.com/library/fbfae092-93b1-b72f-6b42-a498a1543e00%28Office.15%29.aspx)|
|[CacheSettings](http://msdn.microsoft.com/library/48b25030-cbb7-2fec-8025-01b8a96bf6eb%28Office.15%29.aspx)|
|[CacheStatus](http://msdn.microsoft.com/library/77d4498f-bc75-7d97-3d12-4edc9263f32e%28Office.15%29.aspx)|
|[CalculateAll](http://msdn.microsoft.com/library/147d5036-6397-7c3c-cff2-2876ea9b3e0f%28Office.15%29.aspx)|
|[CalculateProject](http://msdn.microsoft.com/library/2581daef-d563-1fd2-4540-65cfbf5ae390%28Office.15%29.aspx)|
|[CalendarBarStyles](http://msdn.microsoft.com/library/bf168abd-3033-f187-ee3e-19e672be4aac%28Office.15%29.aspx)|
|[CalendarBarStylesEdit](http://msdn.microsoft.com/library/6ae39422-20bb-dd77-0d0b-0d130dfdbfe5%28Office.15%29.aspx)|
|[CalendarBarStylesEditEx](http://msdn.microsoft.com/library/3b7cb188-fff6-b9c1-a673-34774791c043%28Office.15%29.aspx)|
|[CalendarBestFitWeekHeight](http://msdn.microsoft.com/library/58b7e8e8-4001-ef47-c7ba-71af617768eb%28Office.15%29.aspx)|
|[CalendarDateBoxes](http://msdn.microsoft.com/library/3870fa41-ef58-8b5d-efe1-b8b3d3a03835%28Office.15%29.aspx)|
|[CalendarDateBoxesEx](http://msdn.microsoft.com/library/a6c1fffd-ce21-d3ef-348f-1f41b5231005%28Office.15%29.aspx)|
|[CalendarDateShading](http://msdn.microsoft.com/library/fedb04c6-e9a4-9289-aedd-042f3751e27d%28Office.15%29.aspx)|
|[CalendarDateShadingEdit](http://msdn.microsoft.com/library/73c8875c-fc54-ae8a-55de-f2640ac4c23a%28Office.15%29.aspx)|
|[CalendarDateShadingEditEx](http://msdn.microsoft.com/library/13382dff-e043-480e-a9f7-300d743bd62a%28Office.15%29.aspx)|
|[CalendarLayout](http://msdn.microsoft.com/library/c948c118-c50f-493d-ba3a-e43ee0d50fa3%28Office.15%29.aspx)|
|[CalendarShowBarSplits](http://msdn.microsoft.com/library/d52f7a1e-ec74-3804-4bbd-3e27ae362e26%28Office.15%29.aspx)|
|[CalendarTaskList](http://msdn.microsoft.com/library/dc37a9b6-616b-248d-d597-fcfbe5074ab1%28Office.15%29.aspx)|
|[CalendarTimescale](http://msdn.microsoft.com/library/4a3cbf04-974b-b83b-b552-572b7c48e31b%28Office.15%29.aspx)|
|[CalendarWeekHeadingsEx](http://msdn.microsoft.com/library/af964116-1d0e-7ab8-4674-4418c1c80f9c%28Office.15%29.aspx)|
|[ChangeColumnDataType](http://msdn.microsoft.com/library/25cbcb73-4cbd-3ea7-ff16-90a4d3028af9%28Office.15%29.aspx)|
|[ChangeStatusDate](http://msdn.microsoft.com/library/93635ef2-43c2-7cfd-5869-f8270a95a0ea%28Office.15%29.aspx)|
|[ChangeWorkingTimeEx](http://msdn.microsoft.com/library/4608fdab-0b39-9918-522a-71d502ba7e3a%28Office.15%29.aspx)|
|[CheckField](http://msdn.microsoft.com/library/a3360541-faa7-169e-1b23-5b3937fc6c07%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/dd2cc86f-44f5-9c7e-c4d1-8475d11367ac%28Office.15%29.aspx)|
|[CheckOut](http://msdn.microsoft.com/library/36e19455-a77d-46d5-c5c0-60f07feeba13%28Office.15%29.aspx)|
|[CheckResourceErrors](http://msdn.microsoft.com/library/780cf9c8-078b-3707-f0e4-a468432c1ced%28Office.15%29.aspx)|
|[CheckTaskErrors](http://msdn.microsoft.com/library/7b361295-993a-13b2-b9bb-26f149e16e72%28Office.15%29.aspx)|
|[CleanupCache](http://msdn.microsoft.com/library/cabd3c0b-b4d0-65ee-0fbd-8be2bde3e170%28Office.15%29.aspx)|
|[CleanupProjectFromCache](http://msdn.microsoft.com/library/40fef64a-036f-8e1c-ce86-0c3609777f77%28Office.15%29.aspx)|
|[ClearConstraint](http://msdn.microsoft.com/library/7a6e9e98-0f0d-6fdd-61b2-c13cdb0cbd7a%28Office.15%29.aspx)|
|[CloseComparison](http://msdn.microsoft.com/library/27c4dc50-7a85-fe92-f294-e5d568b88ed2%28Office.15%29.aspx)|
|[CloseUndoTransaction](http://msdn.microsoft.com/library/704bde43-803d-fd63-68a6-7b4058e5d3b1%28Office.15%29.aspx)|
|[ColumnAlignment](http://msdn.microsoft.com/library/9c51eb2d-c28b-cb00-57e5-1643093e4acb%28Office.15%29.aspx)|
|[ColumnBestFit](http://msdn.microsoft.com/library/51f96761-33ab-d2e3-7a1e-c8266bdaa7a1%28Office.15%29.aspx)|
|[ColumnDelete](http://msdn.microsoft.com/library/a492d8ab-6ed6-49f8-e626-d0a042546021%28Office.15%29.aspx)|
|[ColumnEdit](http://msdn.microsoft.com/library/16fbcb23-419f-9e25-9f3b-271b0d5eda3d%28Office.15%29.aspx)|
|[ColumnInsert](http://msdn.microsoft.com/library/5dfa6b58-7d13-4a96-fdea-8cbe95af52eb%28Office.15%29.aspx)|
|[ComAddInsDialog](http://msdn.microsoft.com/library/06889c2c-2c3a-355d-34c9-ca1d3c31ed2b%28Office.15%29.aspx)|
|[CommitmentsPane](http://msdn.microsoft.com/library/5b37e396-7c70-4554-8164-ea05406ed299%28Office.15%29.aspx)|
|[CompareProjectsLegendToggle](http://msdn.microsoft.com/library/a43d9ff8-9384-5189-ffdc-ac139e791779%28Office.15%29.aspx)|
|[CompareProjectVersions](http://msdn.microsoft.com/library/82af9450-0cec-f7b4-df5c-81ecea3b662f%28Office.15%29.aspx)|
|[ConsolidateProjects](http://msdn.microsoft.com/library/6f1f719c-09c0-076f-4680-24ac26a6538d%28Office.15%29.aspx)|
|[ConvertHangulToHanja](http://msdn.microsoft.com/library/0617dd57-1e0e-a54d-1739-c92efac25237%28Office.15%29.aspx)|
|[CopyReport](http://msdn.microsoft.com/library/9f1e59d5-a2a5-4c8f-1c01-b1c63046558d%28Office.15%29.aspx)|
|[CreateComparisonReport](http://msdn.microsoft.com/library/55b423a7-4613-e1ba-c1b8-e790e74694e7%28Office.15%29.aspx)|
|[CreateEnterpriseCalendar](http://msdn.microsoft.com/library/5d53083b-f34e-d604-6d77-b232eea0eb71%28Office.15%29.aspx)|
|[CreateProjectSite](http://msdn.microsoft.com/library/79c77f3c-0ea6-eed7-762c-f364dc7f3ab7%28Office.15%29.aspx)|
|[CustomFieldDelete](http://msdn.microsoft.com/library/8778f6ee-61bb-b4d0-8846-8c16717cd494%28Office.15%29.aspx)|
|[CustomFieldGetFormula](http://msdn.microsoft.com/library/ce741a1a-1227-b3ae-f45e-0d1f3a048311%28Office.15%29.aspx)|
|[CustomFieldGetName](http://msdn.microsoft.com/library/c68a6aae-7350-e4b5-318b-3d11b77847de%28Office.15%29.aspx)|
|[CustomFieldIndicatorAdd](http://msdn.microsoft.com/library/dc5d071b-3cf8-fe56-df16-c5a6051142da%28Office.15%29.aspx)|
|[CustomFieldIndicatorDelete](http://msdn.microsoft.com/library/729eafe9-4d1a-07a6-efbc-ab0c94e3af59%28Office.15%29.aspx)|
|[CustomFieldIndicators](http://msdn.microsoft.com/library/afbb7bff-49fe-7e12-a257-cab4c730bfbb%28Office.15%29.aspx)|
|[CustomFieldMappingDialog](http://msdn.microsoft.com/library/cb4bd820-04c0-7364-4fde-3a1f4534b72e%28Office.15%29.aspx)|
|[CustomFieldPropertiesEx](http://msdn.microsoft.com/library/3eac9820-848a-011a-96df-f752ea33f31f%28Office.15%29.aspx)|
|[CustomFieldRename](http://msdn.microsoft.com/library/0ca77914-1881-eee5-a8ec-7b47c6464969%28Office.15%29.aspx)|
|[CustomFieldSetFormula](http://msdn.microsoft.com/library/d6d5a5d5-c948-07c9-3f5e-b4607df6538c%28Office.15%29.aspx)|
|[CustomFieldValueList](http://msdn.microsoft.com/library/7365511c-6746-869b-f8e7-d4b87c5b8e70%28Office.15%29.aspx)|
|[CustomFieldValueListAdd](http://msdn.microsoft.com/library/6ef6c528-dc7a-00e8-a770-70b3b9ab86ae%28Office.15%29.aspx)|
|[CustomFieldValueListDelete](http://msdn.microsoft.com/library/f8c513b6-2aab-3e42-ca97-7f91f88f5b61%28Office.15%29.aspx)|
|[CustomFieldValueListGetItem](http://msdn.microsoft.com/library/54ab8b15-374a-3c7a-ffe6-bc90b5d4561e%28Office.15%29.aspx)|
|[CustomForms](http://msdn.microsoft.com/library/392bdcf3-59af-cfa4-c14f-a5d7a6f07495%28Office.15%29.aspx)|
|[CustomizeField](http://msdn.microsoft.com/library/e02fef90-4dc0-639e-d06e-65db997baa8e%28Office.15%29.aspx)|
|[CustomizeIMEMode](http://msdn.microsoft.com/library/1e6cae3d-7b06-327a-4db1-8b4416d703ee%28Office.15%29.aspx)|
|[CustomOutlineCodeEditEx](http://msdn.microsoft.com/library/fc0f60a6-18bf-a8e6-9376-1222a126a64a%28Office.15%29.aspx)|
|[DateAdd](http://msdn.microsoft.com/library/df0da054-495c-c224-ebc8-b47acb78e2af%28Office.15%29.aspx)|
|[DateDifference](http://msdn.microsoft.com/library/7f34e866-5cd3-971d-42ee-39e7768c1273%28Office.15%29.aspx)|
|[DateFormat](http://msdn.microsoft.com/library/b4fc14a0-5139-b7cf-8d96-443cd23fd8ec%28Office.15%29.aspx)|
|[DateSubtract](http://msdn.microsoft.com/library/1eb05a59-271d-31d0-8945-23bc3c9600e0%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/307b1373-309a-1ecf-6899-fd64e663e4f9%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/a517c66f-4bec-9bec-270c-2053bc733145%28Office.15%29.aspx)|
|[DDELinksUpdate](http://msdn.microsoft.com/library/590b5379-f9b7-b245-beed-f656eadd8269%28Office.15%29.aspx)|
|[DDEPasteLink](http://msdn.microsoft.com/library/f97547e7-b541-1a77-94a4-96da1a52ecb2%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/92753522-dad8-4312-eef0-49fd075cea3f%28Office.15%29.aspx)|
|[DeleteFromDatabase](http://msdn.microsoft.com/library/22bed2ff-0e8b-e589-1479-06c482f296a9%28Office.15%29.aspx)|
|[DependenciesPane](http://msdn.microsoft.com/library/c4365a73-af82-7074-9a3e-51298c2dcff6%28Office.15%29.aspx)|
|[DetailsPaneToggle](http://msdn.microsoft.com/library/f62a42b2-397f-45c0-f2c1-f0468b8d489b%28Office.15%29.aspx)|
|[DetailStylesAdd](http://msdn.microsoft.com/library/40a1dfa4-ef57-835d-4e42-9631c906ac0b%28Office.15%29.aspx)|
|[DetailStylesFormat](http://msdn.microsoft.com/library/df3b7963-134f-be55-715e-2e4c214b35fc%28Office.15%29.aspx)|
|[DetailStylesFormatEx](http://msdn.microsoft.com/library/3e460e76-ff7b-f07b-058c-1e37c53e453e%28Office.15%29.aspx)|
|[DetailStylesProperties](http://msdn.microsoft.com/library/f066f826-eef2-7f97-dafa-998f7bd70f42%28Office.15%29.aspx)|
|[DetailStylesRemove](http://msdn.microsoft.com/library/67be5a7d-f066-f22c-7df1-834caeb7b6e2%28Office.15%29.aspx)|
|[DetailStylesRemoveAll](http://msdn.microsoft.com/library/71e9a154-3c02-f289-a06b-b1bbe74f7f70%28Office.15%29.aspx)|
|[DetailStylesToggleItem](http://msdn.microsoft.com/library/744022ac-e5c1-ee5a-c02b-c6962c821c55%28Office.15%29.aspx)|
|[DisplaySharedWorkspace](http://msdn.microsoft.com/library/6d2b53de-8375-75e8-4d1a-2516464de1ce%28Office.15%29.aspx)|
|[DistributeTableColumns](http://msdn.microsoft.com/library/e8523495-e90b-4a01-5c99-c522dd140704%28Office.15%29.aspx)|
|[DistributeTableRows](http://msdn.microsoft.com/library/6a8b9fcf-4922-52ae-d4f9-306d22692224%28Office.15%29.aspx)|
|[DocClose](http://msdn.microsoft.com/library/ddcd72c1-11e7-aa15-12da-ef26d3545742%28Office.15%29.aspx)|
|[DocMaximize](http://msdn.microsoft.com/library/8a24ca4e-e39d-ddae-869d-02d928c27393%28Office.15%29.aspx)|
|[DocMove](http://msdn.microsoft.com/library/defa6ea7-5d1a-d3c4-6486-39192d1da99c%28Office.15%29.aspx)|
|[DocRestore](http://msdn.microsoft.com/library/78589202-af87-2ab9-d03e-93fb48067481%28Office.15%29.aspx)|
|[DocSize](http://msdn.microsoft.com/library/03eb42ef-748e-ef42-a453-8305b0e2835c%28Office.15%29.aspx)|
|[DocumentExport](http://msdn.microsoft.com/library/891bf868-1256-2688-cdb2-2bccfbf2afc2%28Office.15%29.aspx)|
|[DocumentLibraryVersionsDialog](http://msdn.microsoft.com/library/650b9b22-91e0-c565-16c3-b7b72c8bb473%28Office.15%29.aspx)|
|[DrawingCreate](http://msdn.microsoft.com/library/fc146a90-8207-0708-4cca-2015912b284a%28Office.15%29.aspx)|
|[DrawingCycleColor](http://msdn.microsoft.com/library/2465b550-ff0d-360e-0881-641f23fc61c8%28Office.15%29.aspx)|
|[DrawingMove](http://msdn.microsoft.com/library/0d6e2b43-a9ab-1e9d-ad89-afa01afddb50%28Office.15%29.aspx)|
|[DrawingProperties](http://msdn.microsoft.com/library/8d63be84-6321-c0b2-27f0-945baf349714%28Office.15%29.aspx)|
|[DrawingReshape](http://msdn.microsoft.com/library/b9fe0b7c-4112-92fd-d66b-3ebe64e75b8d%28Office.15%29.aspx)|
|[DurationFormat](http://msdn.microsoft.com/library/37970edc-c6f9-66b7-7c0d-b22beb8a36c1%28Office.15%29.aspx)|
|[DurationValue](http://msdn.microsoft.com/library/745acbd3-600c-1179-1d61-be0dab88cdf5%28Office.15%29.aspx)|
|[EditClear](http://msdn.microsoft.com/library/0f87ca1c-c87c-774a-e8dd-2f4d29a40e28%28Office.15%29.aspx)|
|[EditClearFormats](http://msdn.microsoft.com/library/3d8ad4e8-5f3f-80e8-821d-dc44a842d982%28Office.15%29.aspx)|
|[EditClearHyperlink](http://msdn.microsoft.com/library/386e9e73-5c65-0baf-2125-4dbb50675eb1%28Office.15%29.aspx)|
|[EditCopy](http://msdn.microsoft.com/library/a3c1ed1a-d865-80bc-df42-8e0165b4f158%28Office.15%29.aspx)|
|[EditCopyPicture](http://msdn.microsoft.com/library/03f6306b-3538-9a34-dbc3-4ff2f7f40b1e%28Office.15%29.aspx)|
|[EditCut](http://msdn.microsoft.com/library/63b43184-4dcf-d863-87a9-af93c54d4001%28Office.15%29.aspx)|
|[EditDelete](http://msdn.microsoft.com/library/db224f69-ac74-5c5d-6547-7df93ac54eab%28Office.15%29.aspx)|
|[EditEnterpriseCalendar](http://msdn.microsoft.com/library/f40f98f4-82cc-6576-c41e-a9bdd5adb9b8%28Office.15%29.aspx)|
|[EditGoTo](http://msdn.microsoft.com/library/cd2c886b-fddf-d7b8-8f16-51a3af5f0005%28Office.15%29.aspx)|
|[EditHyperlink](http://msdn.microsoft.com/library/d652ccc4-207e-933f-c281-a2d5d7db0b76%28Office.15%29.aspx)|
|[EditInsert](http://msdn.microsoft.com/library/c6b858c7-beb4-9e0e-b3f3-39a1fc37d106%28Office.15%29.aspx)|
|[EditPaste](http://msdn.microsoft.com/library/88413d66-9ccb-99c4-35ca-f6b51d984e22%28Office.15%29.aspx)|
|[EditPasteAsHyperlink](http://msdn.microsoft.com/library/7a2c31dc-43a4-0dc0-2d5c-ee4de18263e4%28Office.15%29.aspx)|
|[EditPasteSpecial](http://msdn.microsoft.com/library/afbe96f1-a4f6-e879-cacc-115761f5e1c4%28Office.15%29.aspx)|
|[EditRedo](http://msdn.microsoft.com/library/4d391a2e-cc0b-f2c6-2347-8020ada46670%28Office.15%29.aspx)|
|[EditTPStyle](http://msdn.microsoft.com/library/71252516-31b5-1184-97f8-da27558620f1%28Office.15%29.aspx)|
|[EditUndo](http://msdn.microsoft.com/library/f13ce3a1-f8f2-8b00-d870-6e30f6b772f5%28Office.15%29.aspx)|
|[EnterpriseGlobalCheckOut](http://msdn.microsoft.com/library/f84fd1bb-1576-8b5b-4d90-4332d0819a6c%28Office.15%29.aspx)|
|[EnterpriseMakeServerURLTrusted](http://msdn.microsoft.com/library/c91df8a2-370c-5f56-c6b4-44239d613ba6%28Office.15%29.aspx)|
|[EnterpriseProjectDelete](http://msdn.microsoft.com/library/ef6c296e-c9d2-02ad-77d1-557c59419872%28Office.15%29.aspx)|
|[EnterpriseProjectImportWizard](http://msdn.microsoft.com/library/0666657f-4352-d7d3-5651-88dc584ea917%28Office.15%29.aspx)|
|[EnterpriseProjectProfiles](http://msdn.microsoft.com/library/b9f9b381-246b-ffc0-e505-1d33fa349fc7%28Office.15%29.aspx)|
|[EnterpriseResourceGet](http://msdn.microsoft.com/library/c1e29298-7859-28c4-edbf-917acdd8aecd%28Office.15%29.aspx)|
|[EnterpriseResourcesImportEx](http://msdn.microsoft.com/library/58b92ff5-da61-07cc-daca-b56e4270a8a4%28Office.15%29.aspx)|
|[EnterpriseResourcesOpen](http://msdn.microsoft.com/library/343b5391-2a28-043d-8ee9-34c71003126c%28Office.15%29.aspx)|
|[EnterpriseResSubstitutionWizard](http://msdn.microsoft.com/library/627b04ad-0088-5032-4f05-b6dc8cabe436%28Office.15%29.aspx)|
|[EnterpriseTeamBuilder](http://msdn.microsoft.com/library/9c164db0-5542-ec3e-121b-206a38cb3ba9%28Office.15%29.aspx)|
|[FieldConstantToFieldName](http://msdn.microsoft.com/library/b8e55035-64e8-fda5-4ad6-9f5e51a55181%28Office.15%29.aspx)|
|[FieldNameToFieldConstant](http://msdn.microsoft.com/library/0830db06-22a7-3ca5-c9ca-f9efbc360767%28Office.15%29.aspx)|
|[FileCloseAllEx](http://msdn.microsoft.com/library/95c7c89f-cfb0-f881-a31b-70ae951fb3f1%28Office.15%29.aspx)|
|[FileCloseEx](http://msdn.microsoft.com/library/56e6eec6-6031-312b-fba5-50db7b43f0b1%28Office.15%29.aspx)|
|[FileExit](http://msdn.microsoft.com/library/a69bc574-dcc3-3710-c705-0566fcf10235%28Office.15%29.aspx)|
|[FileLoadLast](http://msdn.microsoft.com/library/c775d573-d184-d3ac-ed81-3552cc9b045b%28Office.15%29.aspx)|
|[FileNew](http://msdn.microsoft.com/library/59b5acd1-78dc-9fd2-d672-4cdd6a6005aa%28Office.15%29.aspx)|
|[FileOpenEx](http://msdn.microsoft.com/library/d03c13b0-c12f-1d45-bb80-26711d69a378%28Office.15%29.aspx)|
|[FileOpenOrCreate](http://msdn.microsoft.com/library/dced57e2-158a-c323-cf3d-86c493165fa1%28Office.15%29.aspx)|
|[FileOpenUsingBackstage](http://msdn.microsoft.com/library/8e67d279-cbe6-4cfc-f809-ab83c6298e2f%28Office.15%29.aspx)|
|[FilePageSetup](http://msdn.microsoft.com/library/441d787e-8f0d-34ab-09ee-f1e8b1fa350c%28Office.15%29.aspx)|
|[FilePageSetupCalendar](http://msdn.microsoft.com/library/50f4ab0a-ffb4-2bff-44af-82b674de7c4c%28Office.15%29.aspx)|
|[FilePageSetupCalendarText](http://msdn.microsoft.com/library/279e4f0e-f2fb-0822-bf75-700b365c301d%28Office.15%29.aspx)|
|[FilePageSetupCalendarTextEx](http://msdn.microsoft.com/library/370cfaa4-4a7b-e40e-be9e-d562bf9947d7%28Office.15%29.aspx)|
|[FilePageSetupFooter](http://msdn.microsoft.com/library/0ca38a3a-4004-d32b-5a8a-0a4fdb79b68b%28Office.15%29.aspx)|
|[FilePageSetupHeader](http://msdn.microsoft.com/library/e41ff9e1-d656-14fe-3d81-deef3065d11d%28Office.15%29.aspx)|
|[FilePageSetupLegend](http://msdn.microsoft.com/library/b4118a37-f777-b806-9bb4-3f7e6766eda7%28Office.15%29.aspx)|
|[FilePageSetupLegendEx](http://msdn.microsoft.com/library/5cc6c6c1-2228-9c12-3ba6-fd124852a7aa%28Office.15%29.aspx)|
|[FilePageSetupMargins](http://msdn.microsoft.com/library/c36099a7-4ed2-0f0c-c3bb-9af35c88eb35%28Office.15%29.aspx)|
|[FilePageSetupPage](http://msdn.microsoft.com/library/7c5cf66d-715b-17e1-a03a-a376617a1e02%28Office.15%29.aspx)|
|[FilePageSetupView](http://msdn.microsoft.com/library/46a90db8-a635-3592-77ed-c051afa36946%28Office.15%29.aspx)|
|[FilePrint](http://msdn.microsoft.com/library/47937a14-3c57-a597-0b67-5c095bda8ec7%28Office.15%29.aspx)|
|[FilePrintPreview](http://msdn.microsoft.com/library/b17921eb-0c61-35ed-4cf6-44321f301510%28Office.15%29.aspx)|
|[FilePrintSetup](http://msdn.microsoft.com/library/87c49847-3b00-28d7-f45b-3205947a6627%28Office.15%29.aspx)|
|[FileProperties](http://msdn.microsoft.com/library/e1edf1f2-52e1-8a90-aef8-5a5453e89178%28Office.15%29.aspx)|
|[FileSave](http://msdn.microsoft.com/library/2c0ca58c-98f6-2264-51a8-0c93d10816f9%28Office.15%29.aspx)|
|[FileSaveAs](http://msdn.microsoft.com/library/0b5fe86c-28ea-5a9e-53df-5a83030c0d20%28Office.15%29.aspx)|
|[FileSaveOffline](http://msdn.microsoft.com/library/109f95d5-be49-549f-fa39-3231207d61de%28Office.15%29.aspx)|
|[FileSaveWorkspace](http://msdn.microsoft.com/library/f7c524e5-aa9e-e1a2-6f32-defb7cc23f04%28Office.15%29.aspx)|
|[FillAcross](http://msdn.microsoft.com/library/9ab6a32a-84b4-e9c5-2632-b02205275e82%28Office.15%29.aspx)|
|[FillDown](http://msdn.microsoft.com/library/5ccb5f67-64c1-9230-ca58-52bd9bd2c4d5%28Office.15%29.aspx)|
|[FilterApply](http://msdn.microsoft.com/library/d270862e-0577-a9db-e63b-9dcf1dc68b4a%28Office.15%29.aspx)|
|[FilterClear](http://msdn.microsoft.com/library/5de6ac7d-79c5-15e3-5d10-cbf8dd0ccde7%28Office.15%29.aspx)|
|[FilterEdit](http://msdn.microsoft.com/library/e576d3e2-5ac9-006a-2151-dc918b71eef8%28Office.15%29.aspx)|
|[FilterNew](http://msdn.microsoft.com/library/9289cf4f-ce29-695d-baf8-08316ed1e31b%28Office.15%29.aspx)|
|[Filters](http://msdn.microsoft.com/library/f192d400-9867-b978-c68f-e4bc262d36c7%28Office.15%29.aspx)|
|[FilterShowSummaryRows](http://msdn.microsoft.com/library/173bf591-7579-505f-3cbd-42eaddb231ad%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/0e7b1027-5609-19fa-f100-4eb7b108bae7%28Office.15%29.aspx)|
|[FindEx](http://msdn.microsoft.com/library/fdb2661e-f705-ffa4-1ca3-7bbc97b9958d%28Office.15%29.aspx)|
|[FindFile](http://msdn.microsoft.com/library/2f420df9-f234-4990-70cb-5891780e0359%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/005d4cf9-0262-b485-348c-9feb4d7ab389%28Office.15%29.aspx)|
|[FindPrevious](http://msdn.microsoft.com/library/424d20d6-ecec-f46c-62b1-b44f40a40043%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/d612e80b-93c1-7312-d164-be552b580370%28Office.15%29.aspx)|
|[Font32Ex](http://msdn.microsoft.com/library/5f4928a6-d7b3-ff30-48ef-a5037dbeff21%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/6bd38edc-a2af-d6d5-8e46-898b700135b2%28Office.15%29.aspx)|
|[FontEx](http://msdn.microsoft.com/library/4904d4b1-dacb-8020-0c4e-3af0503c68ba%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/d5e79d03-af96-98fb-8f80-6c1fa583a215%28Office.15%29.aspx)|
|[FontStrikethrough](http://msdn.microsoft.com/library/e8689bfe-1c74-5582-8bf1-97b089207321%28Office.15%29.aspx)|
|[FontUnderLine](http://msdn.microsoft.com/library/a093b42b-6b4a-b775-ad81-f85cb940ab88%28Office.15%29.aspx)|
|[Form](http://msdn.microsoft.com/library/23e7c800-bda9-c931-bc27-084dec872953%28Office.15%29.aspx)|
|[FormatCopy](http://msdn.microsoft.com/library/d67082ab-01f5-df2c-377d-c539b3863ef0%28Office.15%29.aspx)|
|[FormatPainter](http://msdn.microsoft.com/library/fb2e2fa1-2e14-26ea-6057-583871e4b170%28Office.15%29.aspx)|
|[FormatPaste](http://msdn.microsoft.com/library/605d0f1d-8a4c-955b-7f82-6c84ad98fbef%28Office.15%29.aspx)|
|[FormViewShow](http://msdn.microsoft.com/library/c1e40d2a-a4bd-60af-3e3c-146e97d7e770%28Office.15%29.aspx)|
|[GanttBarEditEx](http://msdn.microsoft.com/library/b574b975-a869-31ba-e525-df8775330b0a%28Office.15%29.aspx)|
|[GanttBarFormat](http://msdn.microsoft.com/library/2b3b3933-1993-d4cf-f4ff-475c4b003514%28Office.15%29.aspx)|
|[GanttBarFormatEx](http://msdn.microsoft.com/library/9ec9d5a3-7cbb-bfed-9571-e6ba657aaeef%28Office.15%29.aspx)|
|[GanttBarLinks](http://msdn.microsoft.com/library/80f8fdaa-e08f-3c5e-64dc-43d3dccd7f86%28Office.15%29.aspx)|
|[GanttBarSize](http://msdn.microsoft.com/library/691ee987-a62b-bf5f-0088-0f153aa64966%28Office.15%29.aspx)|
|[GanttBarStyleBaseline](http://msdn.microsoft.com/library/c9cb0ebb-998c-c9ea-9d3f-5cb06813c364%28Office.15%29.aspx)|
|[GanttBarStyleCritical](http://msdn.microsoft.com/library/2db96bf5-2a33-2894-8fcb-dcb4842bba4c%28Office.15%29.aspx)|
|[GanttBarStyleDelete](http://msdn.microsoft.com/library/3cac2b37-147c-f1bf-bc94-d2bc9bffa14b%28Office.15%29.aspx)|
|[GanttBarStyleEdit](http://msdn.microsoft.com/library/a955c65c-5579-bd76-150e-d98b5045302d%28Office.15%29.aspx)|
|[GanttBarStyleLate](http://msdn.microsoft.com/library/824760ce-0692-de6a-cf50-90307d94f82a%28Office.15%29.aspx)|
|[GanttBarStyleSlack](http://msdn.microsoft.com/library/ccd8feb0-8551-c3fd-3ce5-ca90baaff910%28Office.15%29.aspx)|
|[GanttBarStyleSlippage](http://msdn.microsoft.com/library/2c5ec6cd-d588-a43a-7b06-8338ecd8ae6e%28Office.15%29.aspx)|
|[GanttBarTextDateFormat](http://msdn.microsoft.com/library/b6159c2a-2e4d-dbfc-53dc-040e1ba6cf7a%28Office.15%29.aspx)|
|[GanttChartWizard](http://msdn.microsoft.com/library/e174c0ac-3f31-a98f-a9ad-11a6785c5052%28Office.15%29.aspx)|
|[GanttRollup](http://msdn.microsoft.com/library/8bb5ef38-d0c7-7425-a6ac-e50c7ae979d8%28Office.15%29.aspx)|
|[GanttShowBarSplits](http://msdn.microsoft.com/library/6f3cf637-4718-8fb9-aed9-cd45ef785ca8%28Office.15%29.aspx)|
|[GanttShowDrawings](http://msdn.microsoft.com/library/8e18c9f0-f434-6aea-f6e6-13263011812a%28Office.15%29.aspx)|
|[GetCellInfo](http://msdn.microsoft.com/library/ddd531b1-e66d-5c70-c4ed-2e2b456e3a3b%28Office.15%29.aspx)|
|[GetCurrentTheme](http://msdn.microsoft.com/library/42384278-abaa-c15a-953f-b1ab4d0901c1%28Office.15%29.aspx)|
|[GetProjectServerSettingsEx](http://msdn.microsoft.com/library/cd630197-60e0-0ba8-e01e-114b82fe9f1e%28Office.15%29.aspx)|
|[GetProjectServerVersion](http://msdn.microsoft.com/library/f41cb738-3a30-f555-9d10-78343fae0ddb%28Office.15%29.aspx)|
|[GetRedoListCount](http://msdn.microsoft.com/library/c505545d-4dda-7b0e-42c2-46591e711b74%28Office.15%29.aspx)|
|[GetRedoListItem](http://msdn.microsoft.com/library/65a23a84-dc85-2935-c673-87643d1a2a2d%28Office.15%29.aspx)|
|[GetThemedColor](http://msdn.microsoft.com/library/d7d464cd-a6d0-72b9-33cd-d5d9e7f30b80%28Office.15%29.aspx)|
|[GetUndoListCount](http://msdn.microsoft.com/library/f152c08c-293a-edd4-5d72-49ba1178715c%28Office.15%29.aspx)|
|[GetUndoListItem](http://msdn.microsoft.com/library/e77826ab-118d-2b69-6f99-cb8ce65afb43%28Office.15%29.aspx)|
|[GoalAreaChange](http://msdn.microsoft.com/library/84341db8-3f8e-44f3-4b34-e702ee2841dd%28Office.15%29.aspx)|
|[GoalAreaHighlight](http://msdn.microsoft.com/library/56146d8b-f986-0ba7-3661-26b508db3ec8%28Office.15%29.aspx)|
|[GoalAreaTaskHighlight](http://msdn.microsoft.com/library/32616617-d34a-c9f4-8ddd-17fa3f1c7e74%28Office.15%29.aspx)|
|[GoToItemInVersions](http://msdn.microsoft.com/library/51b7e580-978d-17cc-f293-bb30d77c48c2%28Office.15%29.aspx)|
|[GotoNextOverAllocation](http://msdn.microsoft.com/library/ebe227a1-cd4c-778e-90be-bd2c65c38c95%28Office.15%29.aspx)|
|[GotoTaskDates](http://msdn.microsoft.com/library/d9d3de8d-e4d7-89f4-0dcf-be132287e19e%28Office.15%29.aspx)|
|[Gridlines](http://msdn.microsoft.com/library/36252fa9-e0de-f221-58fb-871c1ddb2f77%28Office.15%29.aspx)|
|[GridlinesEdit](http://msdn.microsoft.com/library/75b9d660-88b5-da71-faf8-215abce897d2%28Office.15%29.aspx)|
|[GridlinesEditEx](http://msdn.microsoft.com/library/fad3c4cc-2643-4af1-ca6b-f376b24a97bb%28Office.15%29.aspx)|
|[GroupApply](http://msdn.microsoft.com/library/862ff123-2fef-611a-f7c3-dedf8eab0e0b%28Office.15%29.aspx)|
|[GroupBy](http://msdn.microsoft.com/library/3756b876-c67c-966f-7df2-f6a129d404f8%28Office.15%29.aspx)|
|[GroupClear](http://msdn.microsoft.com/library/f30532b6-6fe6-afed-2b38-279d8fbb82eb%28Office.15%29.aspx)|
|[GroupMaintainHierarchy](http://msdn.microsoft.com/library/63f5763a-0ca3-d25b-06ac-03e52cdcf6e2%28Office.15%29.aspx)|
|[GroupNew](http://msdn.microsoft.com/library/28db77c8-209a-9833-eb52-f77c23e6dc8c%28Office.15%29.aspx)|
|[Groups](http://msdn.microsoft.com/library/28a1a91f-16e8-16de-9d8b-baee6d67c840%28Office.15%29.aspx)|
|[HelpAbout](http://msdn.microsoft.com/library/8afda354-3914-37f6-c274-bdb816477506%28Office.15%29.aspx)|
|[HelpAnswerWizard](http://msdn.microsoft.com/library/d23eca0c-2145-e6b8-da1c-924169cf01ee%28Office.15%29.aspx)|
|[HelpContents](http://msdn.microsoft.com/library/f45cfb9f-b482-c70d-85cc-bd2936e4ab7d%28Office.15%29.aspx)|
|[HelpLaunch](http://msdn.microsoft.com/library/05e4e98c-bda7-5b41-372b-2f3752d2ab0e%28Office.15%29.aspx)|
|[HelpTechnicalSupport](http://msdn.microsoft.com/library/bbc15d5b-ef91-3899-3ae2-cce5fbb3d328%28Office.15%29.aspx)|
|[HighlightDrivenSuccessors](http://msdn.microsoft.com/library/2c93505b-541f-15a7-31ff-fcddcfa0bb55%28Office.15%29.aspx)|
|[HighlightDrivingPredecessors](http://msdn.microsoft.com/library/2a2653c5-6b7d-9429-f73f-e65c0cda1c5c%28Office.15%29.aspx)|
|[HighlightPredecessors](http://msdn.microsoft.com/library/e4c51516-2e5d-3ef9-3165-84fe6f9ad38b%28Office.15%29.aspx)|
|[HighlightSuccessors](http://msdn.microsoft.com/library/7a72cc0a-49f0-c95d-23cc-35d7ee077539%28Office.15%29.aspx)|
|[ImportCommitment](http://msdn.microsoft.com/library/ad87bf6a-5409-bd10-b658-b81a3ba501f4%28Office.15%29.aspx)|
|[ImportOutlookTasks](http://msdn.microsoft.com/library/74764d22-eb1b-d7ac-fd63-2151f03e85dc%28Office.15%29.aspx)|
|[InactivateTaskToggle](http://msdn.microsoft.com/library/af937c95-b434-95b8-7ea4-848c25ca30bc%28Office.15%29.aspx)|
|[InformationDialog](http://msdn.microsoft.com/library/644b39d6-be73-5a07-4376-02df25d31a02%28Office.15%29.aspx)|
|[InsertBlankRow](http://msdn.microsoft.com/library/1726e283-d242-53d4-d675-b9cb9d649d29%28Office.15%29.aspx)|
|[InsertHyperlink](http://msdn.microsoft.com/library/d5a6ffc3-8cfe-e6c9-c347-4e3a739f6b1a%28Office.15%29.aspx)|
|[InsertManualTask](http://msdn.microsoft.com/library/4fcfa1be-2a92-9906-2024-6bd14a31fdac%28Office.15%29.aspx)|
|[InsertMilestoneTask](http://msdn.microsoft.com/library/a90ebcc2-b779-0c78-124d-f2c0a9ccd2ca%28Office.15%29.aspx)|
|[InsertNotes](http://msdn.microsoft.com/library/aa57d3c7-31d6-c7b2-7cda-576368a686a1%28Office.15%29.aspx)|
|[InsertResource](http://msdn.microsoft.com/library/e3e62534-3a78-28a2-fb87-ed017b83f9fb%28Office.15%29.aspx)|
|[InsertScheduledTask](http://msdn.microsoft.com/library/0bf89c86-6e0b-19fb-131c-70be563876bd%28Office.15%29.aspx)|
|[InsertSummaryTask](http://msdn.microsoft.com/library/efcbf0d9-5912-d6c4-9204-e939af0193ad%28Office.15%29.aspx)|
|[InsertTask](http://msdn.microsoft.com/library/fe4676bf-8d9a-d6e9-2d5e-74fd047c3944%28Office.15%29.aspx)|
|[IsCommandEnabled](http://msdn.microsoft.com/library/22202fed-7531-0f87-0e38-3ee703717ec1%28Office.15%29.aspx)|
|[IsOfficeTaskPaneVisible](http://msdn.microsoft.com/library/822ad2fd-de35-8340-7b24-56e59fb874b4%28Office.15%29.aspx)|
|[IsOffline](http://msdn.microsoft.com/library/fd844bc5-4b7f-7f4c-a11b-5b26bfe314d2%28Office.15%29.aspx)|
|[IsReducedFunctionalityMode](http://msdn.microsoft.com/library/d53320db-377d-2e78-10b2-03af8d8bded3%28Office.15%29.aspx)|
|[IsUndoingOrRedoing](http://msdn.microsoft.com/library/e0e5ddc7-aa22-0d43-1de6-83a260d57608%28Office.15%29.aspx)|
|[IsURLTrusted](http://msdn.microsoft.com/library/850f5c99-7412-3da7-e136-04f86cd7c42d%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/0b1aaddf-503b-37ff-f9f4-eb102a6ca885%28Office.15%29.aspx)|
|[LayoutNow](http://msdn.microsoft.com/library/8f01c461-a58d-7977-bf62-fda787e9334b%28Office.15%29.aspx)|
|[LayoutRelatedNow](http://msdn.microsoft.com/library/a76cca88-86ad-3fb8-82c6-5ce64a074d54%28Office.15%29.aspx)|
|[LayoutSelectionNow](http://msdn.microsoft.com/library/79d8521a-2760-7e73-f430-f39dc7747cd8%28Office.15%29.aspx)|
|[LevelingClear](http://msdn.microsoft.com/library/fdd537eb-f9c2-c8d9-ec26-0f4af9a63c33%28Office.15%29.aspx)|
|[LevelingOptions](http://msdn.microsoft.com/library/388a2315-e44b-3890-a16a-92ea5a778bbd%28Office.15%29.aspx)|
|[LevelingOptionsEx](http://msdn.microsoft.com/library/f8799750-fecf-48d1-7559-25cd7a8d3d28%28Office.15%29.aspx)|
|[LevelNow](http://msdn.microsoft.com/library/c15b4b91-c005-5f7f-0617-2992a2695e1b%28Office.15%29.aspx)|
|[LevelSelected](http://msdn.microsoft.com/library/1e9383cc-43d3-b479-9b95-cf6fb8cf05b1%28Office.15%29.aspx)|
|[LinksBetweenProjects](http://msdn.microsoft.com/library/63962df8-05ef-f3b4-7ad7-4c75b50ac398%28Office.15%29.aspx)|
|[LinkTasks](http://msdn.microsoft.com/library/cc41c963-533c-97bf-8301-388bb2aaf746%28Office.15%29.aspx)|
|[LinkTasksEdit](http://msdn.microsoft.com/library/51c1d75e-afb6-ae8c-162d-15e24c81bd06%28Office.15%29.aspx)|
|[LinkToTaskList](http://msdn.microsoft.com/library/65ae7bd0-446f-74dd-15fc-0a260342be90%28Office.15%29.aspx)|
|[LoadWebBrowserControlEx](http://msdn.microsoft.com/library/2dca75d3-30ad-ecd0-a465-1190234b9b9b%28Office.15%29.aspx)|
|[LoadWebPaneControl](http://msdn.microsoft.com/library/b807a6e0-5a85-14a0-a87f-e4b6181c9648%28Office.15%29.aspx)|
|[LocaleID](http://msdn.microsoft.com/library/aa84a612-3f7a-b47b-7ddc-39d99b1860e7%28Office.15%29.aspx)|
|[LookUpTableAddEx](http://msdn.microsoft.com/library/5f316f1e-de4b-2fe4-6d3e-84a9944adaed%28Office.15%29.aspx)|
|[Macro](http://msdn.microsoft.com/library/e07686b6-3c38-7413-692b-aac8fb9bf526%28Office.15%29.aspx)|
|[MacroSecurity](http://msdn.microsoft.com/library/5b2fc876-50b2-e30b-ab2b-aa3dc3bddc13%28Office.15%29.aspx)|
|[MacroShowCode](http://msdn.microsoft.com/library/671c557f-0f56-a751-d7bb-37d3c2266687%28Office.15%29.aspx)|
|[MacroShowVba](http://msdn.microsoft.com/library/f585dbe3-0f3a-2552-0770-c395072b6aad%28Office.15%29.aspx)|
|[MailLogoff](http://msdn.microsoft.com/library/e8634331-404c-6e01-4ce9-2dac8dcf364c%28Office.15%29.aspx)|
|[MailLogon](http://msdn.microsoft.com/library/0047a6ea-ea36-498c-e744-c4c88a08baae%28Office.15%29.aspx)|
|[MailPostDocument](http://msdn.microsoft.com/library/568d283a-3765-6371-fb2e-31624f15a0ed%28Office.15%29.aspx)|
|[MailRoutingSlip](http://msdn.microsoft.com/library/1ac860a4-b3fc-9305-5b9f-bf0f8b4ea6e1%28Office.15%29.aspx)|
|[MailSend](http://msdn.microsoft.com/library/250c7eed-2bfa-f80f-13d1-c7ca8d6453d1%28Office.15%29.aspx)|
|[MailSession](http://msdn.microsoft.com/library/00f67414-eb0d-6b2a-d557-26812aaee04c%28Office.15%29.aspx)|
|[MailSystem](http://msdn.microsoft.com/library/4ee9011c-f5f5-d0aa-0cd6-aa90130af4af%28Office.15%29.aspx)|
|[MakeFieldEnterprise](http://msdn.microsoft.com/library/ba9564c9-faa6-bce6-0d59-05dee0cfc887%28Office.15%29.aspx)|
|[MakeLocalCalendarEnterprise](http://msdn.microsoft.com/library/deb355ad-39ca-77cd-7d0d-f5915c7185da%28Office.15%29.aspx)|
|[ManageSiteColumns](http://msdn.microsoft.com/library/1900552c-6320-2ff5-4a07-bc6ebee60696%28Office.15%29.aspx)|
|[MapEdit](http://msdn.microsoft.com/library/316d596e-95b3-d616-c8d6-21da651ff284%28Office.15%29.aspx)|
|[Message](http://msdn.microsoft.com/library/d601b101-5338-f404-e63e-6d1ce926a3d7%28Office.15%29.aspx)|
|[NewTasksStartOn](http://msdn.microsoft.com/library/c5009674-105e-a861-56f0-4847926d6c36%28Office.15%29.aspx)|
|[ObjectChangeIcon](http://msdn.microsoft.com/library/8153748e-9b46-5d57-eaaf-0f09564c55e4%28Office.15%29.aspx)|
|[ObjectConvert](http://msdn.microsoft.com/library/31b7cd47-b592-1425-f2b5-53292306019a%28Office.15%29.aspx)|
|[ObjectInsert](http://msdn.microsoft.com/library/2956dd32-9e28-76e9-c991-12650ee48576%28Office.15%29.aspx)|
|[ObjectLinks](http://msdn.microsoft.com/library/fd83706e-cbdf-fcab-9e64-1867952800f8%28Office.15%29.aspx)|
|[ObjectVerb](http://msdn.microsoft.com/library/55507406-5a36-0361-3b91-7f17860dc577%28Office.15%29.aspx)|
|[OfficeOnTheWeb](http://msdn.microsoft.com/library/ea51e58c-c677-7061-e9a6-8bdfc81779b7%28Office.15%29.aspx)|
|[OfficeTaskPaneHide](http://msdn.microsoft.com/library/51ed3c6b-b938-a128-cb27-8f6c2330963f%28Office.15%29.aspx)|
|[OpenBrowser](http://msdn.microsoft.com/library/92691162-1c5f-43b6-57f2-8d56fa3f7bb6%28Office.15%29.aspx)|
|[OpenFromSharePoint](http://msdn.microsoft.com/library/415f8b11-5c6f-d9df-fb58-61ff7f392b5f%28Office.15%29.aspx)|
|[OpenServerPage](http://msdn.microsoft.com/library/6b7e18fd-2ae1-47a0-45fb-58d6b6e27074%28Office.15%29.aspx)|
|[OpenUndoTransaction](http://msdn.microsoft.com/library/b94b2c87-786c-46d6-50d3-d20614493f8f%28Office.15%29.aspx)|
|[OpenXML](http://msdn.microsoft.com/library/dcf3dd0e-78ec-b95c-b890-dca5507acd92%28Office.15%29.aspx)|
|[OptionsCalculation](http://msdn.microsoft.com/library/608d5bd2-eb6b-0e3c-789a-c376ee55816d%28Office.15%29.aspx)|
|[OptionsCalendar](http://msdn.microsoft.com/library/bde3b645-3417-ee45-57b5-0109bc7b17ad%28Office.15%29.aspx)|
|[OptionsEditEx](http://msdn.microsoft.com/library/d735d118-f004-ba67-7aa5-290ff256da10%28Office.15%29.aspx)|
|[OptionsGeneralEx](http://msdn.microsoft.com/library/c82b09d5-0937-ed06-58d6-e6b5fda186ac%28Office.15%29.aspx)|
|[OptionsInterfaceEx](http://msdn.microsoft.com/library/da4dc69c-021f-7ecb-22f6-aebf1d9252dd%28Office.15%29.aspx)|
|[OptionsSave](http://msdn.microsoft.com/library/658a4b31-8bd6-8dbb-852f-a7f604386215%28Office.15%29.aspx)|
|[OptionsSchedule](http://msdn.microsoft.com/library/24035b34-0364-e830-864a-801150e2668d%28Office.15%29.aspx)|
|[OptionsSecurityEx](http://msdn.microsoft.com/library/9c6e0c77-6873-1a90-fb85-ca33ca7c9ec1%28Office.15%29.aspx)|
|[OptionsSecurityTab](http://msdn.microsoft.com/library/f19ecd9c-2507-e437-7780-cf4998b7fd48%28Office.15%29.aspx)|
|[OptionsSpelling](http://msdn.microsoft.com/library/e0085f68-a57d-c117-cc81-ad11f363c5f4%28Office.15%29.aspx)|
|[OptionsViewEx](http://msdn.microsoft.com/library/88abc2b7-116f-4243-f86f-5f4ad9cf8e72%28Office.15%29.aspx)|
|[Organizer](http://msdn.microsoft.com/library/4269290c-7be9-a0af-526d-bde73114c24b%28Office.15%29.aspx)|
|[OrganizerDeleteItem](http://msdn.microsoft.com/library/7c243672-0e31-e224-eadd-3545f7efcde4%28Office.15%29.aspx)|
|[OrganizerMoveItem](http://msdn.microsoft.com/library/a597c657-130e-2e7b-3837-7e3f95421af7%28Office.15%29.aspx)|
|[OrganizerRenameItem](http://msdn.microsoft.com/library/97ef4b63-a2fb-35ac-0a27-ebe8566fd28c%28Office.15%29.aspx)|
|[OutlineHideSubTasks](http://msdn.microsoft.com/library/79e79b71-aa4d-eb17-7f27-96d4dd382547%28Office.15%29.aspx)|
|[OutlineIndent](http://msdn.microsoft.com/library/43225efc-8b41-5ab3-b646-5012fc9453f4%28Office.15%29.aspx)|
|[OutlineOutdent](http://msdn.microsoft.com/library/4972d60f-4da2-78d1-cbab-28eb9a06a8aa%28Office.15%29.aspx)|
|[OutlineShowAllTasks](http://msdn.microsoft.com/library/b8c089b5-f981-cdfd-7378-9e62259b43b4%28Office.15%29.aspx)|
|[OutlineShowSubTasks](http://msdn.microsoft.com/library/f4a1d5c0-f848-e614-cfe5-0142f88d498d%28Office.15%29.aspx)|
|[OutlineShowTasks](http://msdn.microsoft.com/library/614eb1fc-93eb-3df2-ae52-4fad98c80b3b%28Office.15%29.aspx)|
|[OutlineSymbolsToggle](http://msdn.microsoft.com/library/ea65d093-1a07-7bfc-b8bb-4669f0609ecf%28Office.15%29.aspx)|
|[PageBreakRemove](http://msdn.microsoft.com/library/94c82693-4dd3-d178-06b6-e6f0301aa7e1%28Office.15%29.aspx)|
|[PageBreakSet](http://msdn.microsoft.com/library/0d7b831f-7343-e773-36ef-cedd780f9cc5%28Office.15%29.aspx)|
|[PageBreaksRemoveAll](http://msdn.microsoft.com/library/c3fe7794-e43d-f6f5-a9ec-07326bdfd61d%28Office.15%29.aspx)|
|[PageBreaksShow](http://msdn.microsoft.com/library/320e8ddf-6ded-8f64-0de8-a4cc1275e462%28Office.15%29.aspx)|
|[PaneClose](http://msdn.microsoft.com/library/07a0a80f-f036-db3e-a252-ca70de4cb815%28Office.15%29.aspx)|
|[PaneCreate](http://msdn.microsoft.com/library/6ecf7151-eaeb-4a28-c877-a6e5366e2a8e%28Office.15%29.aspx)|
|[PaneNext](http://msdn.microsoft.com/library/7e8543e4-af6a-82ad-8225-16df72d47492%28Office.15%29.aspx)|
|[PanZoomPanTo](http://msdn.microsoft.com/library/7bdca9f2-d006-6cab-872b-01cf54f6e8ce%28Office.15%29.aspx)|
|[PanZoomZoomTo](http://msdn.microsoft.com/library/bd8510b8-fbdb-2c96-94a7-98c377b2d331%28Office.15%29.aspx)|
|[PasteAsPicture](http://msdn.microsoft.com/library/06b85596-281a-b77d-56d1-8c4283a4dba7%28Office.15%29.aspx)|
|[PasteDestFormatting](http://msdn.microsoft.com/library/4a56bb42-d3d7-fcad-d361-63135e23fc3a%28Office.15%29.aspx)|
|[PasteSourceFormatting](http://msdn.microsoft.com/library/3544cad7-51d4-fd80-5aaa-396fb26a0d17%28Office.15%29.aspx)|
|[ProgressLines](http://msdn.microsoft.com/library/d1c56c86-3882-bfa1-dff8-ed42dd5ce8f2%28Office.15%29.aspx)|
|[ProjectCheckOut](http://msdn.microsoft.com/library/4c6f065f-a853-8f42-e948-be7a76435c0b%28Office.15%29.aspx)|
|[ProjectMove](http://msdn.microsoft.com/library/ba30bd12-a26a-12e5-8cff-df1a34a58df0%28Office.15%29.aspx)|
|[ProjectStatistics](http://msdn.microsoft.com/library/aa3cbba5-5c06-7daf-0b07-035faf72015d%28Office.15%29.aspx)|
|[ProjectSummaryInfo](http://msdn.microsoft.com/library/7275598c-02b1-7e07-ecdb-04fa0a21f41a%28Office.15%29.aspx)|
|[Publish](http://msdn.microsoft.com/library/8605f6c9-8710-0c08-79c8-8dec2bedfe18%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/0aaba635-6d6a-c4a3-fab3-03451659021b%28Office.15%29.aspx)|
|[ReassignSelectedAssns](http://msdn.microsoft.com/library/ab3df7f1-bc36-2b8a-23d7-30ee0387a785%28Office.15%29.aspx)|
|[RecurringTaskInsert](http://msdn.microsoft.com/library/3e993c50-54e3-7373-8459-05706eca72c6%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/25a43bd7-4bfd-2be6-172d-8e5bef781f00%28Office.15%29.aspx)|
|[RegisterProject](http://msdn.microsoft.com/library/66cc4443-2adc-ff66-976e-da52c6d4f7ff%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/5e9305ad-ae42-14e9-8e20-f3068d994200%28Office.15%29.aspx)|
|[RemoveHighlight](http://msdn.microsoft.com/library/334f33a1-8c96-9876-0e71-495336fc947b%28Office.15%29.aspx)|
|[RenameReport](http://msdn.microsoft.com/library/8c4a3ac6-e722-97cb-fe61-c617375c8239%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/fd1c66ba-c611-ec97-ebb9-92ff0739c719%28Office.15%29.aspx)|
|[ReplaceEx](http://msdn.microsoft.com/library/af284688-0701-abc7-4d04-b258957fa9dc%28Office.15%29.aspx)|
|[ReportPrint](http://msdn.microsoft.com/library/4117b555-2985-f129-65aa-9f6804ebf221%28Office.15%29.aspx)|
|[ReportPrintPreview](http://msdn.microsoft.com/library/f93003ee-c25e-9581-191e-478bb30314f0%28Office.15%29.aspx)|
|[Reports](http://msdn.microsoft.com/library/5288cc2d-538f-59c8-6c69-2244b1179cc1%28Office.15%29.aspx)|
|[ReportsDialog](http://msdn.microsoft.com/library/92883d01-10bc-7465-1fe0-aa20ad762257%28Office.15%29.aspx)|
|[RequestProgressInformation](http://msdn.microsoft.com/library/a86ec09d-f9c8-07e3-68f4-898c604c3600%28Office.15%29.aspx)|
|[RescheduleToNextAvailable](http://msdn.microsoft.com/library/4245e739-66f9-b40d-3a13-918028986674%28Office.15%29.aspx)|
|[ResetTPStyle](http://msdn.microsoft.com/library/aba4187b-5af3-3a8d-7486-038e9bdae0ae%28Office.15%29.aspx)|
|[ResourceActiveDirectory](http://msdn.microsoft.com/library/d86f7d15-6ec1-711e-b382-95dd908aee7b%28Office.15%29.aspx)|
|[ResourceAddressBook](http://msdn.microsoft.com/library/012ba9fe-f86e-4d1c-ab24-7a500d8f3b0a%28Office.15%29.aspx)|
|[ResourceAssignment](http://msdn.microsoft.com/library/aceb1802-4b5f-0ad3-bd14-ce77c24705fb%28Office.15%29.aspx)|
|[ResourceAssignmentDialog](http://msdn.microsoft.com/library/efe91944-bdfa-a15c-6f28-44fe4d629974%28Office.15%29.aspx)|
|[ResourceCalendarEditDays](http://msdn.microsoft.com/library/0dc0172f-bc49-347a-7c46-f6a6dc608d8f%28Office.15%29.aspx)|
|[ResourceCalendarReset](http://msdn.microsoft.com/library/3dd5a235-c855-0d65-a664-655c9c1fa7b0%28Office.15%29.aspx)|
|[ResourceCalendars](http://msdn.microsoft.com/library/8c40cfad-ec40-43a4-5698-de5abaea7243%28Office.15%29.aspx)|
|[ResourceComparison](http://msdn.microsoft.com/library/42223a8d-cc71-26c0-35e8-c184b40a46c2%28Office.15%29.aspx)|
|[ResourceDetails](http://msdn.microsoft.com/library/63ac7f3c-38c6-6da9-e442-373da02b63a2%28Office.15%29.aspx)|
|[ResourceGraphBarStyles](http://msdn.microsoft.com/library/b8d2baf3-7025-e330-a582-451ec0d115c0%28Office.15%29.aspx)|
|[ResourceGraphBarStylesEx](http://msdn.microsoft.com/library/903c3894-77c9-bd0a-dee0-85c7fcadea38%28Office.15%29.aspx)|
|[ResourceMappingDialog](http://msdn.microsoft.com/library/b465a823-769f-7e3e-2f2c-98bda2502e0a%28Office.15%29.aspx)|
|[ResourceSharing](http://msdn.microsoft.com/library/c11f9715-83c2-7872-1d53-fb538ed21c74%28Office.15%29.aspx)|
|[ResourceSharingPoolAction](http://msdn.microsoft.com/library/0406765b-b6d7-ad6b-c1c2-51bb55591e69%28Office.15%29.aspx)|
|[ResourceSharingPoolRefresh](http://msdn.microsoft.com/library/8ebb9461-67b6-bfd1-771b-1c7d2d3b79df%28Office.15%29.aspx)|
|[ResourceSharingPoolUpdate](http://msdn.microsoft.com/library/1ebcf06f-fce3-7403-2adb-56f60ab73259%28Office.15%29.aspx)|
|[ResourceWindowsAccount](http://msdn.microsoft.com/library/f03e8445-10a6-d288-b6ae-9ea2eb46f532%28Office.15%29.aspx)|
|[RestoreSheetSelection](http://msdn.microsoft.com/library/cbc4dd00-4055-b505-661b-e2c0276335b3%28Office.15%29.aspx)|
|[RowClear](http://msdn.microsoft.com/library/374b031a-bc06-baf3-51de-79b8df03bd02%28Office.15%29.aspx)|
|[RowDelete](http://msdn.microsoft.com/library/71a512ff-4b2f-971c-2c11-a468b3b7afad%28Office.15%29.aspx)|
|[RowInsert](http://msdn.microsoft.com/library/b9d574b8-8565-9eab-f1a3-4a990bf05bd3%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/0d4060b0-79e8-ad48-f5bf-c1050af379a2%28Office.15%29.aspx)|
|[SaveForSharing](http://msdn.microsoft.com/library/a4f46990-aff1-52da-d1c7-7fd99e85d97a%28Office.15%29.aspx)|
|[SaveSheetSelection](http://msdn.microsoft.com/library/ed792b68-7af2-2b42-9f92-eb77e3b1780e%28Office.15%29.aspx)|
|[SegmentBorderColor](http://msdn.microsoft.com/library/99c2d2ba-f0c5-b462-5801-ac9c7ee75a02%28Office.15%29.aspx)|
|[SegmentFillColor](http://msdn.microsoft.com/library/3f943b8a-47e9-979a-4755-f7b021db6b0e%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/d2003dd7-a6a2-6964-34cb-5331995c7990%28Office.15%29.aspx)|
|[SelectBeginning](http://msdn.microsoft.com/library/4adf20ae-4fd2-818a-da8c-133c08cad7fb%28Office.15%29.aspx)|
|[SelectCell](http://msdn.microsoft.com/library/7177d0bb-6e0e-8885-4f29-51faa34cea8b%28Office.15%29.aspx)|
|[SelectCellDown](http://msdn.microsoft.com/library/78754f19-651b-d614-fa69-5fccd6b3387c%28Office.15%29.aspx)|
|[SelectCellLeft](http://msdn.microsoft.com/library/39bcb2db-cf65-0dc4-2594-9b3c58c4c7c9%28Office.15%29.aspx)|
|[SelectCellRight](http://msdn.microsoft.com/library/3753531a-5459-eb25-a9b9-2e9f748a0518%28Office.15%29.aspx)|
|[SelectCellUp](http://msdn.microsoft.com/library/d2e2aecc-0a05-7dd5-23da-a47ffe161028%28Office.15%29.aspx)|
|[SelectColumn](http://msdn.microsoft.com/library/5bb674e9-253e-355f-a501-d0aeaef56535%28Office.15%29.aspx)|
|[SelectEnd](http://msdn.microsoft.com/library/c1d050e7-739d-8a4f-01da-b8c093836733%28Office.15%29.aspx)|
|[SelectionExtend](http://msdn.microsoft.com/library/cffc56a0-0b25-2afa-427c-840aa2053921%28Office.15%29.aspx)|
|[SelectRange](http://msdn.microsoft.com/library/16b5925e-393b-3d4f-70d4-89213f521485%28Office.15%29.aspx)|
|[SelectResourceCell](http://msdn.microsoft.com/library/3bae94f3-5661-63ef-47a6-12824d5426d0%28Office.15%29.aspx)|
|[SelectResourceColumn](http://msdn.microsoft.com/library/22b9396b-ddec-cfed-311d-a02face0ae2f%28Office.15%29.aspx)|
|[SelectResourceField](http://msdn.microsoft.com/library/6942d5a5-4072-4a95-f2b7-33bf965e302f%28Office.15%29.aspx)|
|[SelectRow](http://msdn.microsoft.com/library/63d31b23-3edb-9cd9-16c5-ac4ca4555a2c%28Office.15%29.aspx)|
|[SelectRowEnd](http://msdn.microsoft.com/library/4aa9b311-46d7-2424-e675-6be0c61248f3%28Office.15%29.aspx)|
|[SelectRowStart](http://msdn.microsoft.com/library/cbb2c5a8-edbb-5d5e-e4ef-5a952db769c3%28Office.15%29.aspx)|
|[SelectSheet](http://msdn.microsoft.com/library/7e156dbf-20c7-7cbd-5f3d-57ca5d241ba5%28Office.15%29.aspx)|
|[SelectTable](http://msdn.microsoft.com/library/8cf26b2d-4021-cf2a-8f0d-d033965f3629%28Office.15%29.aspx)|
|[SelectTaskAssns](http://msdn.microsoft.com/library/80683610-657f-f298-0275-831da215a93a%28Office.15%29.aspx)|
|[SelectTaskCell](http://msdn.microsoft.com/library/824be785-faa8-b274-bc4c-b830f828528d%28Office.15%29.aspx)|
|[SelectTaskColumn](http://msdn.microsoft.com/library/f4269749-de44-d7dd-de74-c95a046411fe%28Office.15%29.aspx)|
|[SelectTaskField](http://msdn.microsoft.com/library/182bfb43-c1ae-32e1-2e93-7cb035e36bd0%28Office.15%29.aspx)|
|[SelectTimescaleRange](http://msdn.microsoft.com/library/16a4bd12-7a60-c172-6a73-c3552b2baf4b%28Office.15%29.aspx)|
|[SelectToEnd](http://msdn.microsoft.com/library/80de4420-5ea8-1bf3-3509-a9c605570e2b%28Office.15%29.aspx)|
|[SelectTPLineHeight](http://msdn.microsoft.com/library/f637032a-ede4-6164-e796-716bf5f556f1%28Office.15%29.aspx)|
|[SelectTPTask](http://msdn.microsoft.com/library/ef27e878-8c80-ad09-157d-f803ec2e7352%28Office.15%29.aspx)|
|[ServiceOptionsDialog](http://msdn.microsoft.com/library/089c6989-4d46-5930-c0d5-ca6c0a66aa21%28Office.15%29.aspx)|
|[SetActiveCell](http://msdn.microsoft.com/library/fcc225b7-98a6-7b3d-ff3b-22392f09920b%28Office.15%29.aspx)|
|[SetAutoFilter](http://msdn.microsoft.com/library/4e4b4d4a-838b-f9b7-e3ab-d7bfa8efce5f%28Office.15%29.aspx)|
|[SetField](http://msdn.microsoft.com/library/9f0670a9-b7e3-0bb6-40fc-0dcae63a3c19%28Office.15%29.aspx)|
|[SetLTRTable](http://msdn.microsoft.com/library/33aee9ba-da55-c83c-a1cf-27b5751c3fdf%28Office.15%29.aspx)|
|[SetMatchingField](http://msdn.microsoft.com/library/fcd57c26-6463-8821-481f-0c38d072118a%28Office.15%29.aspx)|
|[SetResourceField](http://msdn.microsoft.com/library/fbf71bbe-86cc-c53c-a0c3-0df288e2b480%28Office.15%29.aspx)|
|[SetResourceFieldByID](http://msdn.microsoft.com/library/1309ee61-6b66-db45-ed69-b0b3dd9b8dda%28Office.15%29.aspx)|
|[SetRowHeight](http://msdn.microsoft.com/library/bfa4a87b-9e9f-9937-4b9d-a7b26576a5da%28Office.15%29.aspx)|
|[SetRTLTable](http://msdn.microsoft.com/library/92dc18e3-fa84-a4b2-d032-aa32a4e3957d%28Office.15%29.aspx)|
|[SetShowTaskSuggestions](http://msdn.microsoft.com/library/650dd088-9b38-8706-900d-dad7a6ebf4fd%28Office.15%29.aspx)|
|[SetShowTaskWarnings](http://msdn.microsoft.com/library/43ccb666-c61d-e26a-2645-9fa2cb4b3d72%28Office.15%29.aspx)|
|[SetSidepaneStateButton](http://msdn.microsoft.com/library/21603c44-d9f3-96b6-ee42-df17eb58287a%28Office.15%29.aspx)|
|[SetSplitBar](http://msdn.microsoft.com/library/caf26a56-43ad-1714-79e4-cab013a55f3c%28Office.15%29.aspx)|
|[SetTaskField](http://msdn.microsoft.com/library/44e3df27-8924-ecbb-b655-7dab9a51d96f%28Office.15%29.aspx)|
|[SetTaskFieldByID](http://msdn.microsoft.com/library/b4c74d96-d25b-707e-15f1-5e7f05363360%28Office.15%29.aspx)|
|[SetTaskMode](http://msdn.microsoft.com/library/0d800877-9cd9-97e0-6912-6a8d5f596276%28Office.15%29.aspx)|
|[SetTitleRowHeight](http://msdn.microsoft.com/library/7ee0d6db-9fd5-bcd4-e495-14d0a270ed99%28Office.15%29.aspx)|
|[SetTPField](http://msdn.microsoft.com/library/66867c0a-e5a7-9492-463b-0cb955f020df%28Office.15%29.aspx)|
|[ShareProjectOnline](http://msdn.microsoft.com/library/7742715a-d78a-334b-5655-7047efd28890%28Office.15%29.aspx)|
|[ShowAddNewColumn](http://msdn.microsoft.com/library/2f13b46a-da46-453d-1165-f9a1d9b06377%28Office.15%29.aspx)|
|[ShowIgnoredTaskWarnings](http://msdn.microsoft.com/library/77eeb3ef-511d-af17-56c1-aa717fd7d213%28Office.15%29.aspx)|
|[ShowOSFTaskPane](http://msdn.microsoft.com/library/50109216-a0e4-ed18-ea92-e0689f896b86%28Office.15%29.aspx)|
|[ShowReportDataPane](http://msdn.microsoft.com/library/7f0e991a-df7c-9534-45de-50d3839fbac7%28Office.15%29.aspx)|
|[SidepaneTaskChange](http://msdn.microsoft.com/library/277a9242-b098-8f69-44b8-668175867b42%28Office.15%29.aspx)|
|[SidepaneToggle](http://msdn.microsoft.com/library/882c9bef-f150-7128-a506-388dbe39558d%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/996df315-32ae-eac8-75cb-182a95f74879%28Office.15%29.aspx)|
|[SpellCheckField](http://msdn.microsoft.com/library/4c5cc4c9-b947-c237-7f7e-0d703bd34352%28Office.15%29.aspx)|
|[SpellingCheck](http://msdn.microsoft.com/library/e9eea1ad-f2c1-7683-2c09-802a0d33fcec%28Office.15%29.aspx)|
|[SplitTask](http://msdn.microsoft.com/library/490dcca9-66c5-9284-44ff-a92aa30fadf4%28Office.15%29.aspx)|
|[StopWebBrowserControlNavigation](http://msdn.microsoft.com/library/6f3e0fbd-607e-905e-94ef-b34b2187a515%28Office.15%29.aspx)|
|[SummaryResourceAssignmentsRefresh](http://msdn.microsoft.com/library/2f6c2c0d-b039-a613-51c6-3660c98456a1%28Office.15%29.aspx)|
|[SummaryTasksShow](http://msdn.microsoft.com/library/bb533875-6ab5-d803-aadd-555279908985%28Office.15%29.aspx)|
|[SynchronizeWithSite](http://msdn.microsoft.com/library/1bd749d2-fe3f-ee86-dc27-5e39267901bc%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/c00cd0bd-e653-685e-f646-b22f60a6e507%28Office.15%29.aspx)|
|[TableApply](http://msdn.microsoft.com/library/3d335475-a0b7-dd61-1c93-a668a878d347%28Office.15%29.aspx)|
|[TableCopy](http://msdn.microsoft.com/library/90e0a546-2802-5ba7-6b49-086b32051451%28Office.15%29.aspx)|
|[TableEdit](http://msdn.microsoft.com/library/370ab75d-9b99-b4b3-db5f-96697320bc68%28Office.15%29.aspx)|
|[TableEditEx](http://msdn.microsoft.com/library/953cdbf6-24ac-5e39-9c23-ec05ec9e4809%28Office.15%29.aspx)|
|[TableReset](http://msdn.microsoft.com/library/1db786fb-b79d-0404-fe39-4118e10f3cb4%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/ef71a3c7-9851-fe87-7189-24f821c96ea3%28Office.15%29.aspx)|
|[TaskComparison](http://msdn.microsoft.com/library/61d0c322-39a3-f731-3662-f6cf6709bb12%28Office.15%29.aspx)|
|[TaskDeliverableCreate](http://msdn.microsoft.com/library/61bd8608-8a5f-3555-b769-5ee951f8ebd7%28Office.15%29.aspx)|
|[TaskDeliverableSync](http://msdn.microsoft.com/library/e5903c42-bade-959b-3c20-d02e3cf56b24%28Office.15%29.aspx)|
|[TaskDependencySync](http://msdn.microsoft.com/library/4b6ed7a4-9bde-0600-3715-fc3d25501a5a%28Office.15%29.aspx)|
|[TaskDrivers](http://msdn.microsoft.com/library/5c5e7563-e994-809b-7a9c-34f6ea338241%28Office.15%29.aspx)|
|[TaskInspector](http://msdn.microsoft.com/library/cc2f34af-a4e0-8ad4-5dd1-9cf9663e342b%28Office.15%29.aspx)|
|[TaskMove](http://msdn.microsoft.com/library/7a847c59-b07c-6bf2-90a3-b62d0d080cc6%28Office.15%29.aspx)|
|[TaskMoveToStatusDate](http://msdn.microsoft.com/library/100ec970-ca52-2ac8-f367-c346c40e4c61%28Office.15%29.aspx)|
|[TaskOnTimeline](http://msdn.microsoft.com/library/8201380b-f0ae-4e53-7461-e323ad6fe5e2%28Office.15%29.aspx)|
|[TaskRespectLinks](http://msdn.microsoft.com/library/1910b74a-7ea7-d0eb-97b9-aa79330952a0%28Office.15%29.aspx)|
|[TextStyles32Ex](http://msdn.microsoft.com/library/8e1ed2bb-dac4-42d7-616b-a67984dcffa4%28Office.15%29.aspx)|
|[TextStylesEx](http://msdn.microsoft.com/library/674c16c8-8ba5-604f-494c-3b59017e1207%28Office.15%29.aspx)|
|[TimelineExport](http://msdn.microsoft.com/library/a2829e86-5b83-0076-33a3-4c10040ffc17%28Office.15%29.aspx)|
|[TimelineFormat](http://msdn.microsoft.com/library/96f936a1-15be-8df4-4683-cd876c8a69ce%28Office.15%29.aspx)|
|[TimelineGotoSelectedTask](http://msdn.microsoft.com/library/62353aab-b850-bcf9-1d16-c7c794643318%28Office.15%29.aspx)|
|[TimelineInsertTask](http://msdn.microsoft.com/library/4a1833a4-ddbb-577d-fe58-5907644fd127%28Office.15%29.aspx)|
|[TimelineShowHide](http://msdn.microsoft.com/library/237052c0-445b-db78-9a74-10e8742a493d%28Office.15%29.aspx)|
|[TimelineTextOnBar](http://msdn.microsoft.com/library/d57ec0d8-8e35-b6eb-1932-454210bc7dad%28Office.15%29.aspx)|
|[TimelineViewToggle](http://msdn.microsoft.com/library/c5623da2-dd27-c22e-0021-b139e8875401%28Office.15%29.aspx)|
|[Timescale](http://msdn.microsoft.com/library/9e67ec39-030b-5f47-3096-282a03b517d4%28Office.15%29.aspx)|
|[TimescaleEdit](http://msdn.microsoft.com/library/7f1ee80d-8de3-ebde-9961-105a31c62653%28Office.15%29.aspx)|
|[TimescaleNonWorking](http://msdn.microsoft.com/library/bc43da1f-1854-d1ca-f44b-48f660f9336f%28Office.15%29.aspx)|
|[TimescaleNonWorkingEx](http://msdn.microsoft.com/library/50c1b96a-a91c-d538-07b7-44b048c8052b%28Office.15%29.aspx)|
|[ToggleAssignments](http://msdn.microsoft.com/library/1bed946e-d45b-4fa2-7e0d-8602c9197093%28Office.15%29.aspx)|
|[ToggleChangeHighlighting](http://msdn.microsoft.com/library/1b18eb3a-b614-a135-6a82-328cf33c5db8%28Office.15%29.aspx)|
|[TogglePreventResOveralloc](http://msdn.microsoft.com/library/7b6686ab-58c6-e1de-cbb1-618495d5c8ba%28Office.15%29.aspx)|
|[ToggleResourceDetails](http://msdn.microsoft.com/library/b8fe41db-b808-cf3d-2ee9-36afca3cd269%28Office.15%29.aspx)|
|[ToggleTaskDetails](http://msdn.microsoft.com/library/c27dffe7-6814-85f5-9c49-21e0efb12cd1%28Office.15%29.aspx)|
|[ToggleTPAutoExpand](http://msdn.microsoft.com/library/17520aa8-b364-22be-cdc3-62850e77a228%28Office.15%29.aspx)|
|[ToggleTPResourceExpand](http://msdn.microsoft.com/library/a4e39a14-3ba7-25b0-470e-a49c5586d490%28Office.15%29.aspx)|
|[ToggleTPUnassigned](http://msdn.microsoft.com/library/7d9231ac-977e-d86c-c8c3-1aa13b13d7d8%28Office.15%29.aspx)|
|[ToggleTPUnscheduled](http://msdn.microsoft.com/library/f2a44cc5-b11f-f22d-4856-f91d5f67d1c0%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/50e1b5ba-fe4b-d53d-5712-8e2023eb2755%28Office.15%29.aspx)|
|[UndoClear](http://msdn.microsoft.com/library/ee46aa2e-e04a-420f-54aa-76fd4ec5c6c8%28Office.15%29.aspx)|
|[UnlinkTasks](http://msdn.microsoft.com/library/76fefb0b-c137-ac6f-a95e-7950803d561f%28Office.15%29.aspx)|
|[UnloadWebBrowserControl](http://msdn.microsoft.com/library/beccb5ae-102c-4c68-595b-47ff08da72ab%28Office.15%29.aspx)|
|[UpdateFromProjectServer](http://msdn.microsoft.com/library/f37bb573-2d25-b4f9-21ba-109db75962f6%28Office.15%29.aspx)|
|[UpdateProject](http://msdn.microsoft.com/library/a6f80334-7faf-ca95-b5ed-0a9fba516169%28Office.15%29.aspx)|
|[UpdateTasks](http://msdn.microsoft.com/library/4a04e459-9f5c-f944-d39f-dcbbfc48fdab%28Office.15%29.aspx)|
|[UsageViewEntryEx](http://msdn.microsoft.com/library/2aac9824-ab5c-006d-99d2-07e019e6409d%28Office.15%29.aspx)|
|[ViewApply](http://msdn.microsoft.com/library/3e0d3fbd-5aa7-ceb8-b926-79646986d464%28Office.15%29.aspx)|
|[ViewApplyEx](http://msdn.microsoft.com/library/437ec3b5-d42d-ed79-e8c7-220f797023b5%28Office.15%29.aspx)|
|[ViewBar](http://msdn.microsoft.com/library/c1bb0168-4ba9-82c2-8043-ece0138e3695%28Office.15%29.aspx)|
|[ViewCopy](http://msdn.microsoft.com/library/b1ed6b3e-ad95-15f4-80bd-054d608ef9a1%28Office.15%29.aspx)|
|[ViewEditCombination](http://msdn.microsoft.com/library/f5d49a1d-7ead-e704-7be2-8d06e54e221f%28Office.15%29.aspx)|
|[ViewEditSingle](http://msdn.microsoft.com/library/445977e9-e540-14b3-a179-ea132491265e%28Office.15%29.aspx)|
|[ViewReset](http://msdn.microsoft.com/library/ea972480-6417-55a7-9b8e-6cc9944df6c9%28Office.15%29.aspx)|
|[Views](http://msdn.microsoft.com/library/76f29c4c-1854-e136-2d72-d50fe786c26b%28Office.15%29.aspx)|
|[ViewsEx](http://msdn.microsoft.com/library/42567343-54df-fbf2-64a3-79ba72d12866%28Office.15%29.aspx)|
|[ViewShowCost](http://msdn.microsoft.com/library/37f4ca8b-f544-281d-6870-360bc763a400%28Office.15%29.aspx)|
|[ViewShowCumulativeCost](http://msdn.microsoft.com/library/46374294-f37b-a71e-ff17-fb3bdf68928d%28Office.15%29.aspx)|
|[ViewShowCumulativeWork](http://msdn.microsoft.com/library/ca31034e-5080-2e88-5742-b8def3b11278%28Office.15%29.aspx)|
|[ViewShowNotes](http://msdn.microsoft.com/library/6721aa38-185d-4b10-abf3-d7587cd793b5%28Office.15%29.aspx)|
|[ViewShowObjects](http://msdn.microsoft.com/library/2bbe735e-b024-5f28-18bd-ef8335995ca2%28Office.15%29.aspx)|
|[ViewShowOverallocation](http://msdn.microsoft.com/library/e8389cd8-6312-e7a1-ac90-e0c52139695c%28Office.15%29.aspx)|
|[ViewShowPeakUnits](http://msdn.microsoft.com/library/d2027dc0-f763-1e26-c036-d6cc130072c5%28Office.15%29.aspx)|
|[ViewShowPercentAllocation](http://msdn.microsoft.com/library/41da8198-1899-f9af-2ddd-7a992a3c3465%28Office.15%29.aspx)|
|[ViewShowPredecessorsSuccessors](http://msdn.microsoft.com/library/14c92bb3-0e0a-35ac-c587-6b7c75146ff0%28Office.15%29.aspx)|
|[ViewShowRemainingAvailability](http://msdn.microsoft.com/library/9e76e3e1-1f50-d744-3804-70d4ce9cff33%28Office.15%29.aspx)|
|[ViewShowResourcesPredecessors](http://msdn.microsoft.com/library/3f7d0a36-cc1b-f3f2-8e25-d6b898d19afe%28Office.15%29.aspx)|
|[ViewShowResourcesSuccessors](http://msdn.microsoft.com/library/632893a7-70ec-6cd5-56c6-82b216f09d48%28Office.15%29.aspx)|
|[ViewShowSchedule](http://msdn.microsoft.com/library/13788fb3-f3ef-cfdc-e66f-ba67273dd5c9%28Office.15%29.aspx)|
|[ViewShowUnitAvailability](http://msdn.microsoft.com/library/900af4b4-dd2d-483e-b207-6d199c51092b%28Office.15%29.aspx)|
|[ViewShowWork](http://msdn.microsoft.com/library/fc2071b1-9aed-015a-a9b5-67de2a9ae12f%28Office.15%29.aspx)|
|[ViewShowWorkAvailability](http://msdn.microsoft.com/library/909fbc1a-fe49-8121-c103-e287d10a49fa%28Office.15%29.aspx)|
|[VisualReports](http://msdn.microsoft.com/library/4934cdcf-06b0-020c-3741-4ef70944cf98%28Office.15%29.aspx)|
|[VisualReportsEdit](http://msdn.microsoft.com/library/ba439985-f18b-f9a3-23d5-3d5ae39c50dc%28Office.15%29.aspx)|
|[VisualReportsNewTemplate](http://msdn.microsoft.com/library/46fbe1f2-a79a-a0e2-ccfb-2c02ed46b184%28Office.15%29.aspx)|
|[VisualReportsSaveCube](http://msdn.microsoft.com/library/51b65e15-7ab5-79ff-9513-c47b204c1751%28Office.15%29.aspx)|
|[VisualReportsSaveDatabase](http://msdn.microsoft.com/library/edcbaff5-beb1-ba11-fb65-ec26a24ab23d%28Office.15%29.aspx)|
|[VisualReportsView](http://msdn.microsoft.com/library/80742129-71eb-355d-1bb8-f64579eef344%28Office.15%29.aspx)|
|[WBSCodeMaskEdit](http://msdn.microsoft.com/library/37ade035-5235-54ab-92fa-962c4172dcdc%28Office.15%29.aspx)|
|[WBSCodeRenumber](http://msdn.microsoft.com/library/c71f6dd3-5ea5-de60-7cd5-09134fa5a278%28Office.15%29.aspx)|
|[WebAddToFavorites](http://msdn.microsoft.com/library/3cf8b3e7-4dbf-8555-1662-2412e7d420b0%28Office.15%29.aspx)|
|[WebCopyHyperlink](http://msdn.microsoft.com/library/9e08c278-71dd-7cf2-d515-1af6ebf184d4%28Office.15%29.aspx)|
|[WebGoBack](http://msdn.microsoft.com/library/bbc0d3bb-9074-eab6-a65a-58d095bf125f%28Office.15%29.aspx)|
|[WebGoForward](http://msdn.microsoft.com/library/2692d709-58e3-cf21-2dce-f056e6144c7e%28Office.15%29.aspx)|
|[WebHideToolbars](http://msdn.microsoft.com/library/c6e323c9-b1a4-79bb-d714-b7ddaebbf619%28Office.15%29.aspx)|
|[WebOpenFavorites](http://msdn.microsoft.com/library/cb32f74e-ceba-0651-1b17-a61e6bce1bf8%28Office.15%29.aspx)|
|[WebOpenHyperlink](http://msdn.microsoft.com/library/f1da5d5f-45a1-02e0-8783-7f919578e3fe%28Office.15%29.aspx)|
|[WebOpenSearchPage](http://msdn.microsoft.com/library/61db85dc-5773-57f6-d716-7c0e1db6d86b%28Office.15%29.aspx)|
|[WebOpenStartPage](http://msdn.microsoft.com/library/7d043964-8be2-fbf2-7d6c-6ad0454e05cb%28Office.15%29.aspx)|
|[WebRefresh](http://msdn.microsoft.com/library/ef36cbc0-4d11-2328-a2d0-24370f4143c8%28Office.15%29.aspx)|
|[WebSetSearchPage](http://msdn.microsoft.com/library/57d23181-92ae-2f45-a2c4-20059a085e8b%28Office.15%29.aspx)|
|[WebSetStartPage](http://msdn.microsoft.com/library/2ffe7e71-fbdc-e6bc-8eae-9da23e5f63f5%28Office.15%29.aspx)|
|[WebStopLoading](http://msdn.microsoft.com/library/e76165ff-0636-3dff-b525-0ff56f24a38c%28Office.15%29.aspx)|
|[WebToolbar](http://msdn.microsoft.com/library/ff0f557f-ec63-0acd-da89-bc06c857524d%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/8b9b39f8-39e5-b162-d8d9-de9838f7b39e%28Office.15%29.aspx)|
|[WindowArrangeAll](http://msdn.microsoft.com/library/504db965-27ea-d0f5-5830-927555ac801c%28Office.15%29.aspx)|
|[WindowHide](http://msdn.microsoft.com/library/37219d9d-1e50-3341-7618-9827d077d4d8%28Office.15%29.aspx)|
|[WindowMoreWindows](http://msdn.microsoft.com/library/66c50a0c-624d-485b-d6c8-3046643dcb36%28Office.15%29.aspx)|
|[WindowNewWindow](http://msdn.microsoft.com/library/fe0c2bcb-7bee-3bec-9c47-3015938ae75d%28Office.15%29.aspx)|
|[WindowNext](http://msdn.microsoft.com/library/10b5306d-038a-1b1c-9dec-8dd9d8b05dc3%28Office.15%29.aspx)|
|[WindowPrev](http://msdn.microsoft.com/library/f95cf733-fc5c-e454-55b6-11f704dee431%28Office.15%29.aspx)|
|[WindowSplit](http://msdn.microsoft.com/library/cbdea999-4692-a10d-80e3-ae6b4407eebc%28Office.15%29.aspx)|
|[WindowUnhide](http://msdn.microsoft.com/library/438693a7-5b99-e373-6d28-9a42dfcda7d1%28Office.15%29.aspx)|
|[WorkOffline](http://msdn.microsoft.com/library/65a38e80-f311-eb19-359a-da9f1022be71%28Office.15%29.aspx)|
|[WrapText](http://msdn.microsoft.com/library/0aaabac2-ee1d-694c-45ac-f522a0034724%28Office.15%29.aspx)|
|[Zoom](http://msdn.microsoft.com/library/0ac9b17a-6791-31a2-11d4-6d97ade57989%28Office.15%29.aspx)|
|[ZoomCalendar](http://msdn.microsoft.com/library/fc02c827-11a0-380b-9e05-b4452246ff05%28Office.15%29.aspx)|
|[ZoomIn](http://msdn.microsoft.com/library/0a6abf44-68ee-b146-d760-a7f0e1e79d76%28Office.15%29.aspx)|
|[ZoomOut](http://msdn.microsoft.com/library/d72dae84-638c-76c7-2de1-4b02b0a0a731%28Office.15%29.aspx)|
|[ZoomReport](http://msdn.microsoft.com/library/05a0ec6e-1329-2545-df89-5d87af88a454%28Office.15%29.aspx)|
|[ZoomTimescale](http://msdn.microsoft.com/library/d20b2c8a-bef2-5456-73f1-a6fa417b427e%28Office.15%29.aspx)|
|[AddEngagement](http://msdn.microsoft.com/library/61fbd902-1fa1-d591-5618-697e5dc9338d%28Office.15%29.aspx)|
|[EngagementInfo](http://msdn.microsoft.com/library/4e95d901-77a0-f1f7-b754-aefeb720e5ea%28Office.15%29.aspx)|
|[GetDpiScaleFactor](http://msdn.microsoft.com/library/d1e7f1e5-095c-aa4c-0550-1a077c1a2de3%28Office.15%29.aspx)|
|[InsertTimelineBar](http://msdn.microsoft.com/library/2cb9d639-3363-79e3-ced6-73b0a574986a%28Office.15%29.aspx)|
|[Inspector](http://msdn.microsoft.com/library/f386160f-232a-7e4d-37e0-9c090a58df8a%28Office.15%29.aspx)|
|[LocaleName](http://msdn.microsoft.com/library/989d8c73-3452-2abe-fbaa-f68d532e353e%28Office.15%29.aspx)|
|[ProjectSummaryInfoEx](http://msdn.microsoft.com/library/2827f735-6a7b-9f33-c1c6-2c5f1f7492f6%28Office.15%29.aspx)|
|[RefreshEngagementsForProject](http://msdn.microsoft.com/library/f0530b2b-18de-70b8-d27d-a51ded376fe3%28Office.15%29.aspx)|
|[RemoveTimelineBar](http://msdn.microsoft.com/library/8385d889-b81e-5422-a032-c7073fa7c65d%28Office.15%29.aspx)|
|[SubmitAllEngagementsForProject](http://msdn.microsoft.com/library/7e695f9f-5c0b-bbbf-9abe-a695e72591a1%28Office.15%29.aspx)|
|[SubmitSelectedEngagementsForProject](http://msdn.microsoft.com/library/bfa4d8b5-5806-54d9-009e-ff8fcb96d994%28Office.15%29.aspx)|
|[TaskOnTimelineEx](http://msdn.microsoft.com/library/4307f842-0ccc-d7ac-f386-ec8d259011c6%28Office.15%29.aspx)|
|[TimelineBarDateRange](http://msdn.microsoft.com/library/a1d257f3-92b7-6719-4ce5-5b959823e702%28Office.15%29.aspx)|
|[UpdateEngagementsForProject](http://msdn.microsoft.com/library/cda633ec-2143-0f6e-80eb-2d9751d8782f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveCell](http://msdn.microsoft.com/library/880931d8-fc23-7938-e4fe-bd800eeae318%28Office.15%29.aspx)|
|[ActiveProject](http://msdn.microsoft.com/library/07844166-ca9b-15eb-a5e2-6f00a7c0a030%28Office.15%29.aspx)|
|[ActiveSelection](http://msdn.microsoft.com/library/aa72b337-4031-a970-0921-d1d60f66096e%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/57ea4398-b496-96a9-bb5e-4f529f9a5c1e%28Office.15%29.aspx)|
|[AMText](http://msdn.microsoft.com/library/92a8d781-79ac-ebfa-8419-31cbd140e505%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/9e0d273b-21b7-b15f-a269-db8d40d47c72%28Office.15%29.aspx)|
|[AskToUpdateLinks](http://msdn.microsoft.com/library/669aacdc-3e0a-031b-0fea-2becd7aab67f%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/f53bf107-9fd1-78f9-f8db-0b8c2acc5f72%28Office.15%29.aspx)|
|[AutoClearLeveling](http://msdn.microsoft.com/library/799384b3-8d85-b07c-14e3-3d25d7ec3d33%28Office.15%29.aspx)|
|[AutoLevel](http://msdn.microsoft.com/library/dc4fbd05-0493-7699-eb39-ea2af8fddde1%28Office.15%29.aspx)|
|[AutomaticallyFillPhoneticFields](http://msdn.microsoft.com/library/2c4eef7e-bde4-6aa9-b383-7634447997a0%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/08f71d7f-37bf-c845-89c3-a69e34892efe%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/26a8b2d9-0af9-9ec6-ed02-e52229214ce1%28Office.15%29.aspx)|
|[Calculation](http://msdn.microsoft.com/library/eca7ce92-38ad-7bbf-78d2-e06cd3e35b6e%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/e43c55ea-d239-a6e5-42ce-35da5b47aa01%28Office.15%29.aspx)|
|[CellDragAndDrop](http://msdn.microsoft.com/library/a9ce116c-bf06-126b-2955-20e5a2880633%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/32bf64b2-4fee-cc9f-210e-4a463d04a900%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/80f57057-9bb3-018b-0e45-fd1423368091%28Office.15%29.aspx)|
|[CompareProjectsCurrentVersionName](http://msdn.microsoft.com/library/1cd33b06-9c68-7278-9d78-0308f9277e88%28Office.15%29.aspx)|
|[CompareProjectsPreviousVersionName](http://msdn.microsoft.com/library/205c43cc-1dbf-d8ef-5dea-90087d7820ed%28Office.15%29.aspx)|
|[DateOrder](http://msdn.microsoft.com/library/9eba39c8-6e4a-3b8c-69c3-82e078269cda%28Office.15%29.aspx)|
|[DateSeparator](http://msdn.microsoft.com/library/ff89ed80-4839-4195-09a7-f1d6ab4ad88a%28Office.15%29.aspx)|
|[DayLeadingZero](http://msdn.microsoft.com/library/63220c29-6f41-7a32-22bd-0afe49fef5c3%28Office.15%29.aspx)|
|[DecimalSeparator](http://msdn.microsoft.com/library/c331d9fa-c389-16d7-b09b-1a17bba5b3c0%28Office.15%29.aspx)|
|[DefaultAutoFilter](http://msdn.microsoft.com/library/ef2301d0-6a57-7d88-75ee-6b57909317e9%28Office.15%29.aspx)|
|[DefaultDateFormat](http://msdn.microsoft.com/library/01f20463-2d23-0e65-ab54-cc23673509da%28Office.15%29.aspx)|
|[DefaultView](http://msdn.microsoft.com/library/19f3cc23-6267-0b1f-7db5-7783d6936533%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/ef5234ee-cdee-3eee-ca31-1f680d34f9c6%28Office.15%29.aspx)|
|[DisplayEntryBar](http://msdn.microsoft.com/library/56121152-2302-9d32-3a64-68b8b68f0f90%28Office.15%29.aspx)|
|[DisplayOLEIndicator](http://msdn.microsoft.com/library/85d58ecf-69eb-a1c4-c5a2-6499bfa56e22%28Office.15%29.aspx)|
|[DisplayPlanningWizard](http://msdn.microsoft.com/library/eac1ac6f-8d2d-6c4a-fe7c-fadab773a624%28Office.15%29.aspx)|
|[DisplayProjectGuide](http://msdn.microsoft.com/library/5b10db18-8cee-3824-79c7-85eadf11b0af%28Office.15%29.aspx)|
|[DisplayRecentFiles](http://msdn.microsoft.com/library/99c60109-676f-41ee-3ed0-76d0b0c4ee99%28Office.15%29.aspx)|
|[DisplayScheduleMessages](http://msdn.microsoft.com/library/a65e0a34-da09-c57d-d155-eecabcc24922%28Office.15%29.aspx)|
|[DisplayScrollBars](http://msdn.microsoft.com/library/4c8e2aa3-3d85-94c8-d1ce-67586b78e7e7%28Office.15%29.aspx)|
|[DisplayStatusBar](http://msdn.microsoft.com/library/9764173e-6ea3-29d1-5b79-fb763986584b%28Office.15%29.aspx)|
|[DisplayViewBar](http://msdn.microsoft.com/library/e097b5ef-9d87-a55b-719b-3c31c6000b05%28Office.15%29.aspx)|
|[DisplayWindowsInTaskbar](http://msdn.microsoft.com/library/f4b352f4-4b7b-a438-c29b-bc2f5b68aeb0%28Office.15%29.aspx)|
|[DisplayWizardErrors](http://msdn.microsoft.com/library/b0af54ec-392f-b84d-3dcc-cc52c991b66d%28Office.15%29.aspx)|
|[DisplayWizardScheduling](http://msdn.microsoft.com/library/abcd5660-1eef-d53b-548f-6ead0c57f836%28Office.15%29.aspx)|
|[DisplayWizardUsage](http://msdn.microsoft.com/library/3b4362ca-c748-3da8-0e1d-8d0baa1c3d69%28Office.15%29.aspx)|
|[Edition](http://msdn.microsoft.com/library/3277932e-5d23-a5c3-8928-e41557d542e2%28Office.15%29.aspx)|
|[EnableCancelKey](http://msdn.microsoft.com/library/9b5f4f90-3ef3-139b-5f76-f48d3d7710a8%28Office.15%29.aspx)|
|[EnableChangeHighlighting](http://msdn.microsoft.com/library/68365e16-6746-9ee6-9462-f9b076f986c6%28Office.15%29.aspx)|
|[EnterpriseAllowLocalBaseCalendars](http://msdn.microsoft.com/library/91c15501-a321-47fb-7c9a-ebe894ead50a%28Office.15%29.aspx)|
|[EnterpriseListSeparator](http://msdn.microsoft.com/library/973201dd-0c1c-88d5-052a-94028584f6d5%28Office.15%29.aspx)|
|[EnterpriseProtectActuals](http://msdn.microsoft.com/library/99880223-194c-39de-aed0-068b3eb0a96b%28Office.15%29.aspx)|
|[FileBuildID](http://msdn.microsoft.com/library/6fae0673-614d-6cb2-31c2-bff9eabeecc9%28Office.15%29.aspx)|
|[FileFormatID](http://msdn.microsoft.com/library/86a6a5ce-6508-f1ad-b9cc-fb86fd96e410%28Office.15%29.aspx)|
|[GetCacheStatusForProject](http://msdn.microsoft.com/library/71ab8ee0-83fc-c80f-3583-ce66b167d044%28Office.15%29.aspx)|
|[GlobalBaseCalendars](http://msdn.microsoft.com/library/98a498f9-e040-9b00-e84a-806a8a17a181%28Office.15%29.aspx)|
|[GlobalOutlineCodes](http://msdn.microsoft.com/library/a63d1a87-5c87-a2d6-c4da-70ab9526eaae%28Office.15%29.aspx)|
|[GlobalReports](http://msdn.microsoft.com/library/736be78c-2571-b07f-369c-845a06f9d1f9%28Office.15%29.aspx)|
|[GlobalResourceFilters](http://msdn.microsoft.com/library/d3cd1f3f-7d46-612f-eaa1-3b3528ca4ab6%28Office.15%29.aspx)|
|[GlobalResourceTables](http://msdn.microsoft.com/library/8cf96f98-b0d0-2ae8-e472-6f74b62f6411%28Office.15%29.aspx)|
|[GlobalTaskFilters](http://msdn.microsoft.com/library/1f85f0c7-9cb8-e531-c690-6ea795ebaa94%28Office.15%29.aspx)|
|[GlobalTaskTables](http://msdn.microsoft.com/library/5ca768b2-2e0f-6889-a300-8e81130ba798%28Office.15%29.aspx)|
|[GlobalViews](http://msdn.microsoft.com/library/6f85147a-cc5c-dd8a-c091-68af6c3d5c98%28Office.15%29.aspx)|
|[GlobalViewsCombination](http://msdn.microsoft.com/library/9eace5f8-163e-9b55-2ca4-f1bf43bf12d4%28Office.15%29.aspx)|
|[GlobalViewsSingle](http://msdn.microsoft.com/library/5cfb067d-8b8e-7c6c-dca0-286b753f1067%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/e980a85d-218c-b82d-1043-9670cab23560%28Office.15%29.aspx)|
|[IsCheckedOut](http://msdn.microsoft.com/library/616f9342-9d9b-dd85-873c-3e40abfec019%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/5a1b51ca-1621-798d-7bbe-75b565d694fe%28Office.15%29.aspx)|
|[LevelFreeformTasks](http://msdn.microsoft.com/library/d9a9abca-0efa-ea38-3665-7f7b7ecccc9e%28Office.15%29.aspx)|
|[LevelIndividualAssignments](http://msdn.microsoft.com/library/7ce1ac1a-3dd5-be72-f410-7ff173b1c280%28Office.15%29.aspx)|
|[LevelingCanSplit](http://msdn.microsoft.com/library/3c3c523d-5a5f-3b12-f411-97c95793b4c7%28Office.15%29.aspx)|
|[LevelOrder](http://msdn.microsoft.com/library/c8cf70bb-7808-48c4-43b4-c7f693d4613d%28Office.15%29.aspx)|
|[LevelPeriodBasis](http://msdn.microsoft.com/library/24a13a72-8a3d-e59b-d912-6847f79019e1%28Office.15%29.aspx)|
|[LevelProposedBookings](http://msdn.microsoft.com/library/34b1d355-a5c5-38c2-9502-064ecd81906e%28Office.15%29.aspx)|
|[LevelWithinSlack](http://msdn.microsoft.com/library/08c7a6ea-fe7d-c5c5-42b4-66940019aa0b%28Office.15%29.aspx)|
|[ListSeparator](http://msdn.microsoft.com/library/86659bb7-d205-2205-9cd5-e825cdef64ce%28Office.15%29.aspx)|
|[LoadLastFile](http://msdn.microsoft.com/library/2e76f572-d9ad-179a-b32b-b2708898023c%28Office.15%29.aspx)|
|[MonthLeadingZero](http://msdn.microsoft.com/library/b2911e1b-195e-984e-173c-a058a9d3766e%28Office.15%29.aspx)|
|[MoveAfterReturn](http://msdn.microsoft.com/library/03bfce40-c863-a29b-da19-e4c2523265ff%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/a8986bfb-fe80-ff24-cf2a-290c16b3555f%28Office.15%29.aspx)|
|[NewTasksEstimated](http://msdn.microsoft.com/library/cb1fe0c1-7473-e163-104d-2302ffbc8325%28Office.15%29.aspx)|
|[OperatingSystem](http://msdn.microsoft.com/library/0ef34d09-9fc5-ec9e-3d96-416cda925616%28Office.15%29.aspx)|
|[PanZoomFinish](http://msdn.microsoft.com/library/a080b0b7-45fc-7c7e-90ee-7685ac9a1917%28Office.15%29.aspx)|
|[PanZoomStart](http://msdn.microsoft.com/library/7e5ff081-c5fb-165e-8ded-bad1c3cdc72a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4942313c-4f03-362f-0fbb-9596050a7231%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/bb739ed8-9e1f-36e0-5a26-68301cfa24eb%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/6daeb9c9-40e1-1da6-1123-50983dd4d8c2%28Office.15%29.aspx)|
|[PMText](http://msdn.microsoft.com/library/a52193c7-2a74-c3b8-357b-ea7637309d14%28Office.15%29.aspx)|
|[Profiles](http://msdn.microsoft.com/library/4b57eb31-f73d-6587-c555-fc14220e4a2a%28Office.15%29.aspx)|
|[Projects](http://msdn.microsoft.com/library/792b7334-a424-abe1-287e-285d3ab362c7%28Office.15%29.aspx)|
|[PromptForSummaryInfo](http://msdn.microsoft.com/library/c1ce90ec-e52b-397f-640c-4a8da1e17a7f%28Office.15%29.aspx)|
|[RecentFilesMaximum](http://msdn.microsoft.com/library/005c7c09-1fbf-b807-ebe6-601c55e56c97%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/23260017-c550-4f2b-a57f-4d7f7c1c0d52%28Office.15%29.aspx)|
|[ShowAssignmentUnitsAs](http://msdn.microsoft.com/library/bf845895-9efe-bb95-9b60-3fdc30615ab5%28Office.15%29.aspx)|
|[ShowEstimatedDuration](http://msdn.microsoft.com/library/c32670b7-a2e8-a46b-f91d-88b20749fa46%28Office.15%29.aspx)|
|[ShowWelcome](http://msdn.microsoft.com/library/083e38b0-7cfe-027a-882d-05c98f8de3b2%28Office.15%29.aspx)|
|[StartWeekOn](http://msdn.microsoft.com/library/a5e3c262-4450-e6c1-85d7-ca15d324c2aa%28Office.15%29.aspx)|
|[StartYearIn](http://msdn.microsoft.com/library/7662b30f-572d-a7a7-22d1-6a3bb6e1ea5d%28Office.15%29.aspx)|
|[StatusBar](http://msdn.microsoft.com/library/c88965a0-302c-e0ce-ca5b-06fc2d21ff2d%28Office.15%29.aspx)|
|[SupportsMultipleDocuments](http://msdn.microsoft.com/library/d5f1daf1-21b0-3c6c-44b2-8e3f665c7055%28Office.15%29.aspx)|
|[SupportsMultipleWindows](http://msdn.microsoft.com/library/d52eb74c-a809-2084-9e4e-45ca4d53d2e4%28Office.15%29.aspx)|
|[ThousandSeparator](http://msdn.microsoft.com/library/27e0548f-2def-1aa6-6ffb-46fbeba85dca%28Office.15%29.aspx)|
|[TimeLeadingZero](http://msdn.microsoft.com/library/292f06a7-2c3c-f7d7-1577-2b3d06a4731d%28Office.15%29.aspx)|
|[TimescaleFinish](http://msdn.microsoft.com/library/66c07ebc-ee68-bf4c-9af1-c894d4617e44%28Office.15%29.aspx)|
|[TimescaleStart](http://msdn.microsoft.com/library/001e0556-e1b4-d817-868a-834970becc46%28Office.15%29.aspx)|
|[TimeSeparator](http://msdn.microsoft.com/library/e0846c88-f8d6-0c73-d72a-2d0f20ee05ba%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/c6c34d81-5378-2e47-9849-31edf050b4b0%28Office.15%29.aspx)|
|[TrustProjectServerAndWSSPages](http://msdn.microsoft.com/library/c79b17d6-c344-0bed-8087-7f5d5c17d3af%28Office.15%29.aspx)|
|[TwelveHourTimeFormat](http://msdn.microsoft.com/library/899caa96-da4e-8ee6-988a-6cef64a1a46c%28Office.15%29.aspx)|
|[UndoLevels](http://msdn.microsoft.com/library/2cfd6962-2cae-b7fe-2c8d-f0c81a1c1302%28Office.15%29.aspx)|
|[UsableHeight](http://msdn.microsoft.com/library/f0cd8b86-a619-022a-5e26-8d4c5e815af3%28Office.15%29.aspx)|
|[UsableWidth](http://msdn.microsoft.com/library/ccc312da-6794-657d-7c76-e3e8549e2da7%28Office.15%29.aspx)|
|[Use3DLook](http://msdn.microsoft.com/library/df4fce68-5ce1-5a99-3001-597a19871b1c%28Office.15%29.aspx)|
|[UseOMIDs](http://msdn.microsoft.com/library/15339e09-0b65-d939-df47-eb538dee7c38%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/4c67c930-5c15-43cf-7536-ab11661af1a7%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/c501ef16-f4c8-3c08-69b8-3e9756db8336%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/abd71fdd-1ae8-5b29-a2a3-0ffedde3f667%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/9fa235a3-8bdd-a4d3-3d40-e0f77f52e314%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/43bf25de-4908-1fad-e5d5-9fba21e8b03c%28Office.15%29.aspx)|
|[VisualReportsAdditionalTemplatePath](http://msdn.microsoft.com/library/d1727b8c-595e-bf41-cbd5-3cebed893636%28Office.15%29.aspx)|
|[VisualReportTemplateList](http://msdn.microsoft.com/library/b756c00f-7f76-9697-711e-400762cc48c3%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/ee52fc37-ff4e-5e86-77ac-7d60b65397ef%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/0f589af9-d587-3cfc-ffbb-64d901ff3bd4%28Office.15%29.aspx)|
|[Windows2](http://msdn.microsoft.com/library/038d051c-769d-3a14-c884-7b4b669d3cc8%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/1a5d372d-9e05-80b4-6722-19781381d372%28Office.15%29.aspx)|

