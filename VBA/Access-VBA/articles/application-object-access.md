---
title: Application Object (Access)
keywords: vbaac10.chm12627
f1_keywords:
- vbaac10.chm12627
ms.prod: access
api_name:
- Access.Application
ms.assetid: aefb0713-97e6-e2c7-e530-8fd2e1316a55
ms.date: 06/08/2017
---


# Application Object (Access)

The  **Application** object refers to the active Microsoft Access application.


## Remarks

The  **Application** object contains all Access objects and collections.

You can use the  **Application** object to apply methods or property settings to the entire Access application. For example, you can use the **[SetOption](http://msdn.microsoft.com/library/6cb1f036-01c2-16bf-f62a-e5235dfb3c65%28Office.15%29.aspx)** method of the **Application** object to set database options from Visual Basic. The following example shows how you can set the **Display Status Bar** check box on the **Current Database** tab of the **Access Options** dialog box.




```
Application.SetOption "Show Status Bar", True
```

Access is a COM component that supports Automation, formerly called OLE Automation. You can manipulate Access objects from another application that also supports Automation. To do this, you use the  **Application** object.

For example, Microsoft Visual Basic is a COM component. You can open anAccess database from Visual Basic and work with its objects. From Visual Basic, first create a reference to the Access object library. Then create a new instance of the  **Application** class and point an object variable to it, as in the following example:




```
Dim appAccess As New Access.Application
```

From applications that do not support the  **New** keyword, you can create a new instance of the **Application** class by using the **CreateObject** function:




```
Dim appAccess As Object 
Set appAccess = CreateObject("Access.Application")
```

After you create a new instance of the  **Application** class, you can open a database or create a new database, by using either the **[OpenCurrentDatabase](http://msdn.microsoft.com/library/fd214849-02ac-eaa6-7525-9aee42b92f3d%28Office.15%29.aspx)** method or the **[NewCurrentDatabase](http://msdn.microsoft.com/library/6934a77e-5fa0-7e43-e159-2ffc2a944dca%28Office.15%29.aspx)** method. You can then set the properties of the **Application** object and call its methods. When you return a reference to the **CommandBars** object by using the **CommandBars** property of the **Application** object, you can access all Microsoft Office command bar objects and collections by using this reference.

You can also manipulate other Access objects through the  **Application** object. For example, by using the **[OpenForm](http://msdn.microsoft.com/library/a1c9d3a9-2af8-c30a-acb0-6428c70dcdb0%28Office.15%29.aspx)** method of the Access **[DoCmd](docmd-object-access.md)** object, you can open an Access form from Microsoft Office Excel:




```
appAccess.DoCmd.OpenForm "Orders"
```

For more information about creating a reference and controlling objects by using Automation, see the documentation for the application that is acting as the COM component.


## Methods



|**Name**|
|:-----|
|[AccessError](http://msdn.microsoft.com/library/811ef090-bdd4-5d1d-afc5-782470f57483%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/c2024fa1-a972-7798-9bc0-776c6e30c4a4%28Office.15%29.aspx)|
|[BuildCriteria](http://msdn.microsoft.com/library/098e9aca-3dc1-ad21-4374-5d8ae7c80c56%28Office.15%29.aspx)|
|[CloseCurrentDatabase](http://msdn.microsoft.com/library/f5dec73c-54b4-c5ea-7cb9-25b5997f539e%28Office.15%29.aspx)|
|[CodeDb](http://msdn.microsoft.com/library/7f0cff23-1265-231f-9ab5-fa83c19d39cf%28Office.15%29.aspx)|
|[ColumnHistory](http://msdn.microsoft.com/library/e2c1b71f-6561-b38d-8173-9926bc4bd9da%28Office.15%29.aspx)|
|[CompactRepair](http://msdn.microsoft.com/library/4820fd79-d907-21bc-0ad5-5fc096c1ef3b%28Office.15%29.aspx)|
|[ConvertAccessProject](http://msdn.microsoft.com/library/49b865f5-30b6-7b28-efe8-df2cc67951b0%28Office.15%29.aspx)|
|[CreateAccessProject](http://msdn.microsoft.com/library/66628c62-20db-e3a3-5d27-9da3846f0514%28Office.15%29.aspx)|
|[CreateAdditionalData](http://msdn.microsoft.com/library/d27df827-1bcc-eb1e-00d2-46eebd265440%28Office.15%29.aspx)|
|[CreateControl](http://msdn.microsoft.com/library/f5b1689c-62c4-163d-c659-607cee7572f6%28Office.15%29.aspx)|
|[CreateForm](http://msdn.microsoft.com/library/113c8f7f-baf1-bf5c-85ce-6dc1f3d3e942%28Office.15%29.aspx)|
|[CreateGroupLevel](http://msdn.microsoft.com/library/880c1e36-b7b5-7ea4-a2ca-d7c3f0a5a7be%28Office.15%29.aspx)|
|[CreateReport](http://msdn.microsoft.com/library/4b086f8c-8017-0b5f-72a7-7c180c32f52d%28Office.15%29.aspx)|
|[CreateReportControl](http://msdn.microsoft.com/library/4b970377-450b-9909-f5c3-cb7f8445139f%28Office.15%29.aspx)|
|[CurrentDb](http://msdn.microsoft.com/library/defcf58f-7689-90e0-001c-ba5e7e87eb88%28Office.15%29.aspx)|
|[CurrentUser](http://msdn.microsoft.com/library/1cf7ee61-459c-1224-cfdf-a0b051eeb06e%28Office.15%29.aspx)|
|[CurrentWebUser](http://msdn.microsoft.com/library/cb8b230d-71c5-c73d-c88e-1a13246492a5%28Office.15%29.aspx)|
|[CurrentWebUserGroups](http://msdn.microsoft.com/library/efe80f7a-b6ac-12a5-3704-6e662c87e134%28Office.15%29.aspx)|
|[DAvg](http://msdn.microsoft.com/library/966cd884-8693-d1d2-b35b-567e71b7e56d%28Office.15%29.aspx)|
|[DCount](http://msdn.microsoft.com/library/257f0b2a-e23d-2728-afd2-7700b59e5456%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/9828607e-a2e3-15e2-699a-12fb2dc9e897%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/7b05c3ad-574e-d904-5d50-ff646486ef07%28Office.15%29.aspx)|
|[DDEPoke](http://msdn.microsoft.com/library/5f24d625-bd9b-41fd-004c-dccfb0ec41b6%28Office.15%29.aspx)|
|[DDERequest](http://msdn.microsoft.com/library/c6f5f472-aeac-6de9-8133-bebfc5887eee%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/97684f64-dd80-03b6-965d-42e9d0e6f264%28Office.15%29.aspx)|
|[DDETerminateAll](http://msdn.microsoft.com/library/0d2a5e65-c10a-1e78-a0a3-573b9ed804d4%28Office.15%29.aspx)|
|[DefaultWorkspaceClone](http://msdn.microsoft.com/library/f72522e5-dd8d-2cd1-df40-4457ef7f94a6%28Office.15%29.aspx)|
|[DeleteControl](http://msdn.microsoft.com/library/f59f9368-0d7a-8e5f-5140-86e2d2c18c22%28Office.15%29.aspx)|
|[DeleteReportControl](http://msdn.microsoft.com/library/26e30033-ab56-9cfa-3c35-f6d47caf8bd7%28Office.15%29.aspx)|
|[DFirst](http://msdn.microsoft.com/library/670e54ac-a18f-e381-2ca7-257411f92865%28Office.15%29.aspx)|
|[DirtyObject](http://msdn.microsoft.com/library/caf82388-d822-967f-c5f9-0042955ea8d8%28Office.15%29.aspx)|
|[DLast](http://msdn.microsoft.com/library/0a04cbcc-0dbc-4cfc-e5a3-deb9b0f343be%28Office.15%29.aspx)|
|[DLookup](http://msdn.microsoft.com/library/cbe1fc56-e4d7-cb74-02df-48fc379cf432%28Office.15%29.aspx)|
|[DMax](http://msdn.microsoft.com/library/d6d978f2-edad-f478-8c15-bc7aa5b575e0%28Office.15%29.aspx)|
|[DMin](http://msdn.microsoft.com/library/d41b1852-7d97-ddfe-d071-8a1a7b42359b%28Office.15%29.aspx)|
|[DStDev](http://msdn.microsoft.com/library/401b4e16-dfd4-7256-b03d-f3915c5f9ca5%28Office.15%29.aspx)|
|[DStDevP](http://msdn.microsoft.com/library/ca5fb7ad-d91e-1222-e99a-8c55f34482f3%28Office.15%29.aspx)|
|[DSum](http://msdn.microsoft.com/library/53a3cfd4-a5e3-d0c5-1727-070c99d2b984%28Office.15%29.aspx)|
|[DVar](http://msdn.microsoft.com/library/e1566391-4aac-548f-6475-6a8ee63a2bb7%28Office.15%29.aspx)|
|[DVarP](http://msdn.microsoft.com/library/99a2d948-0f38-85fa-6f68-5568262595ae%28Office.15%29.aspx)|
|[Echo](http://msdn.microsoft.com/library/ce94d774-ef06-7cf4-0e91-b5affa41a437%28Office.15%29.aspx)|
|[EuroConvert](http://msdn.microsoft.com/library/35893059-c6cd-d359-f618-94701a50a049%28Office.15%29.aspx)|
|[Eval](http://msdn.microsoft.com/library/d02d5278-1ff3-c405-d579-7a58f2e1ea68%28Office.15%29.aspx)|
|[ExportNavigationPane](http://msdn.microsoft.com/library/49bd679b-d763-ee3e-0cb4-165f1c45f60d%28Office.15%29.aspx)|
|[ExportXML](http://msdn.microsoft.com/library/47627677-d311-c2e1-7532-e8a8a9beef29%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/b5142ca6-8d67-c42b-81a4-5417265a50b0%28Office.15%29.aspx)|
|[GetHiddenAttribute](http://msdn.microsoft.com/library/aee0e022-08d5-10f8-bfd0-588b5310fb43%28Office.15%29.aspx)|
|[GetOption](http://msdn.microsoft.com/library/32736ddf-3551-07f5-1559-d0e139c1697d%28Office.15%29.aspx)|
|[GUIDFromString](http://msdn.microsoft.com/library/943da2f6-a578-f05d-5778-990b6892fc64%28Office.15%29.aspx)|
|[HtmlEncode](http://msdn.microsoft.com/library/294a99f1-9b26-c9ee-0560-8bd54287ebb7%28Office.15%29.aspx)|
|[hWndAccessApp](http://msdn.microsoft.com/library/7a4f162a-e2de-728b-09e0-f9272ad52053%28Office.15%29.aspx)|
|[HyperlinkPart](http://msdn.microsoft.com/library/011665ea-c650-fab3-a736-f26a3de1b65e%28Office.15%29.aspx)|
|[ImportNavigationPane](http://msdn.microsoft.com/library/5365ece3-e2da-031c-4e28-89115d48acf8%28Office.15%29.aspx)|
|[ImportXML](http://msdn.microsoft.com/library/c7baa4be-4ef6-c886-3cd6-06576563b77d%28Office.15%29.aspx)|
|[InstantiateTemplate](http://msdn.microsoft.com/library/de91646c-1681-37e5-30c4-97b42617497b%28Office.15%29.aspx)|
|[IsCurrentWebUserInGroup](http://msdn.microsoft.com/library/49251e19-e375-bcec-29fa-329b2c4fbf3f%28Office.15%29.aspx)|
|[LoadCustomUI](http://msdn.microsoft.com/library/59be6be9-d7a0-98f3-b9c0-57ecba5651f6%28Office.15%29.aspx)|
|[LoadFromAXL](http://msdn.microsoft.com/library/1cce0568-1966-c089-a741-b0934b8676d6%28Office.15%29.aspx)|
|[LoadPicture](http://msdn.microsoft.com/library/d7e64367-c8f2-22c3-6e6e-18eaae9ed07a%28Office.15%29.aspx)|
|[NewAccessProject](http://msdn.microsoft.com/library/e3b3b9ef-31f8-885c-5c92-d269b824fbdb%28Office.15%29.aspx)|
|[NewCurrentDatabase](http://msdn.microsoft.com/library/6934a77e-5fa0-7e43-e159-2ffc2a944dca%28Office.15%29.aspx)|
|[Nz](http://msdn.microsoft.com/library/669fe962-3881-83bb-cc40-ec9b23b44116%28Office.15%29.aspx)|
|[OpenAccessProject](http://msdn.microsoft.com/library/fdc1b231-1512-cbcd-f376-935555861b38%28Office.15%29.aspx)|
|[OpenCurrentDatabase](http://msdn.microsoft.com/library/fd214849-02ac-eaa6-7525-9aee42b92f3d%28Office.15%29.aspx)|
|[PlainText](http://msdn.microsoft.com/library/76a14feb-abee-9306-fe10-27765c4a47c7%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/075ad885-f25d-ea2d-bf74-8ec915265c63%28Office.15%29.aspx)|
|[RefreshDatabaseWindow](http://msdn.microsoft.com/library/63825d35-b24e-ae68-3214-5727dc97eb79%28Office.15%29.aspx)|
|[RefreshTitleBar](http://msdn.microsoft.com/library/9924e3ff-714f-023e-460f-d4aba7702829%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/4cdaf4cb-c25c-aaa4-96ab-52259f9f91c0%28Office.15%29.aspx)|
|[RunCommand](http://msdn.microsoft.com/library/2731352f-7f2d-db3a-314c-e8a789755dd5%28Office.15%29.aspx)|
|[SaveAsAXL](http://msdn.microsoft.com/library/a9557499-7e69-b405-8e2f-d9fcb23fb012%28Office.15%29.aspx)|
|[SaveAsTemplate](http://msdn.microsoft.com/library/3f796181-70c7-f372-92e9-0c2dbbc7262a%28Office.15%29.aspx)|
|[SetDefaultWorkgroupFile](http://msdn.microsoft.com/library/64dc24a0-e6dc-685f-620a-463417e8a25d%28Office.15%29.aspx)|
|[SetHiddenAttribute](http://msdn.microsoft.com/library/b92a1edc-033a-095c-980f-852b8f7e0785%28Office.15%29.aspx)|
|[SetOption](http://msdn.microsoft.com/library/6cb1f036-01c2-16bf-f62a-e5235dfb3c65%28Office.15%29.aspx)|
|[StringFromGUID](http://msdn.microsoft.com/library/527c9459-a62a-9f01-dcda-3c21987b2662%28Office.15%29.aspx)|
|[SysCmd](http://msdn.microsoft.com/library/5064b8cc-6f9a-602b-e304-6d1478d9b4a7%28Office.15%29.aspx)|
|[TransformXML](http://msdn.microsoft.com/library/03b483ad-9785-be26-4632-411d8fc8a19d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/2be2025d-263d-23d9-1b70-fce5108b4875%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/34a5bdb9-8487-49ab-47f1-7c19ace4a633%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/10c259ed-43c2-b413-d137-78b2c9ff4326%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/4589f050-4b0c-8dba-309a-98ad3921baa7%28Office.15%29.aspx)|
|[BrokenReference](http://msdn.microsoft.com/library/20a55f4b-5fe4-9231-bbef-e90c66f88b90%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/d96de996-33f5-a4a1-66d9-c18b3cdbac43%28Office.15%29.aspx)|
|[CodeContextObject](http://msdn.microsoft.com/library/b675d334-33e6-b845-0dd9-6dca36f7b4ab%28Office.15%29.aspx)|
|[CodeData](http://msdn.microsoft.com/library/f75e7676-ec76-9270-109a-91db58e32ff1%28Office.15%29.aspx)|
|[CodeProject](http://msdn.microsoft.com/library/881eeb80-7e78-6ae6-3bb5-e7d67731c48c%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/b94474b4-3690-54ab-1a4b-b30744354db5%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/a7dc2e41-7271-1f2d-b0f9-7fa884311709%28Office.15%29.aspx)|
|[CurrentData](http://msdn.microsoft.com/library/47ddbd6e-cf91-1ccf-e53c-ee999e94d002%28Office.15%29.aspx)|
|[CurrentObjectName](http://msdn.microsoft.com/library/85b32556-96ed-ed3c-dc5b-4c2570639f50%28Office.15%29.aspx)|
|[CurrentObjectType](http://msdn.microsoft.com/library/10065578-b218-8b83-f210-056922a57c4b%28Office.15%29.aspx)|
|[CurrentProject](http://msdn.microsoft.com/library/4efb3378-c1ab-0d60-7617-6df335fcfa03%28Office.15%29.aspx)|
|[DBEngine](http://msdn.microsoft.com/library/ad4638e4-0c72-ce24-e322-e147e2f0cfc2%28Office.15%29.aspx)|
|[DoCmd](http://msdn.microsoft.com/library/171fb56a-b39f-4439-e841-ae4bbbd71719%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/bc9057bc-72a4-0344-a50a-7b73a2d30212%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/8589e1de-e6e7-f85c-0138-0690781d5ed5%28Office.15%29.aspx)|
|[Forms](http://msdn.microsoft.com/library/fbc85a70-538d-b7bf-15e8-c1c7821dc9de%28Office.15%29.aspx)|
|[IsCompiled](http://msdn.microsoft.com/library/c3b80c32-2aba-432c-1909-4c8172a3bebf%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/f2b039bf-95a8-7820-355e-67fa5e47aaf6%28Office.15%29.aspx)|
|[MacroError](http://msdn.microsoft.com/library/08f88f9a-4cb5-850b-a08e-6a2aa62a5bcd%28Office.15%29.aspx)|
|[MenuBar](http://msdn.microsoft.com/library/dc0f6f9c-4627-96a1-83fa-b58ce1eb7236%28Office.15%29.aspx)|
|[Modules](http://msdn.microsoft.com/library/eb99e25f-9a31-82cd-1b61-41c8a227b859%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/63843be1-da9c-8052-52ee-39ca558b5856%28Office.15%29.aspx)|
|[NewFileTaskPane](http://msdn.microsoft.com/library/22b069c2-9c3a-7ee1-e47f-4916a24b32d0%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/ef5e9aef-a0af-b848-638a-df21d0e06963%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/a8398360-f11c-72b9-4b71-7b042889ac9c%28Office.15%29.aspx)|
|[Printers](http://msdn.microsoft.com/library/71383404-8244-6e9b-9c72-8963e0901901%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/b4e374ec-b52f-e73d-174e-bb07f40ab029%28Office.15%29.aspx)|
|[References](http://msdn.microsoft.com/library/da78f26f-1127-796d-bba1-f1c0d98a582e%28Office.15%29.aspx)|
|[Reports](http://msdn.microsoft.com/library/c9fe6b1c-ea14-509e-31f4-dc41f8b99a7f%28Office.15%29.aspx)|
|[ReturnVars](http://msdn.microsoft.com/library/2b8f455a-328f-d2f5-8277-24e9c2b9f5c7%28Office.15%29.aspx)|
|[Screen](http://msdn.microsoft.com/library/d6faa33a-7701-d270-3bc7-04d53ac9303a%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/6785320b-b50f-dcaa-3eae-13d378573613%28Office.15%29.aspx)|
|[TempVars](http://msdn.microsoft.com/library/356f2585-6789-ebe4-5c24-02a361289cd5%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/e82213ac-bd7b-2669-3001-330f40cfdaaa%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/b9ce562e-cfb1-4b39-a287-2c0629f38c7b%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/3fd0113f-5c8f-0477-6030-cf548f7cb2ff%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/ac1558c1-68c4-fdf1-4f59-77343b7b5e59%28Office.15%29.aspx)|
|[WebServices](http://msdn.microsoft.com/library/fed37107-137f-a2c6-96ba-1a97d3c9780a%28Office.15%29.aspx)|

## See also

[Access Object Model Reference](object-model-access-vba-reference.md)


