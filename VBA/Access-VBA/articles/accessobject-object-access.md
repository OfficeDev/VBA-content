---
title: AccessObject Object (Access)
keywords: vbaac10.chm12743
f1_keywords:
- vbaac10.chm12743
ms.prod: access
api_name:
- Access.AccessObject
ms.assetid: 8a770b33-5bff-120a-6707-ca214ee5ced3
ms.date: 06/08/2017
---


# AccessObject Object (Access)

An  **AccessObject** object refers to a particular Access object.


## Remarks

An **AccessObject** object includes information about one instance of an object. The following table list the types of objects each **AccessObject** describes, the name of its collection, and what type of information **AccessObject** contains.



|**AccessObject**|**Collection**|**Contains information about**|
|:-----|:-----|:-----|
|**Database diagram**|**AllDatabaseDiagrams**|Saved database diagrams|
|**Form**|**AllForms**|Saved forms|
|**Function**|**AllFunctions**|Saved functions|
|**Macro**|**AllMacros**|Saved macros|
|**Module**|**AllModules**|Saved modules|
|**Query**|**AllQueries**|Saved queries|
|**Report**|**AllReports**|Saved reports|
|**Stored procedure**|**AllStoredProcedures**|Saved stored procedures|
|**Table**|**AllTables**|Saved tables|
|**View**|**AllViews**|Saved views|
Because an **AccessObject** object corresponds to an existing object, you can't create new **AccessObject** objects or delete existing ones. To refer to an **AccessObject** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:


|Syntax|
|:-----|
|**AllForms** (0)|
|**AllForms** (" _name_ ")|
|**AllForms** ![ _name_ ]|

## Methods



|**Name**|
|:-----|
|[GetDependencyInfo](http://msdn.microsoft.com/library/33feb9c9-abac-cbe4-acf9-989957f41b7a%28Office.15%29.aspx)|
|[IsDependentUpon](http://msdn.microsoft.com/library/aba465c5-4176-c69a-8eb8-1a6737b6d8cf%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[CurrentView](http://msdn.microsoft.com/library/d957f544-9619-be5c-dfce-c6962ba24655%28Office.15%29.aspx)|
|[DateCreated](http://msdn.microsoft.com/library/68a6fd13-2831-386f-0328-274e43219578%28Office.15%29.aspx)|
|[DateModified](http://msdn.microsoft.com/library/a5392776-febe-de09-103d-2d2683f2d0bf%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/9e6d2249-893e-8b0f-87b8-c427e6d89927%28Office.15%29.aspx)|
|[IsLoaded](http://msdn.microsoft.com/library/5e68398c-8a95-f3e1-87ec-e2d637f34429%28Office.15%29.aspx)|
|[IsWeb](http://msdn.microsoft.com/library/57fa0b00-6f1b-b865-a697-b6d3fdd03f82%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/e58b445b-c69e-599a-7396-72a77113e226%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/3db6009b-6c7e-65de-4033-1d592b122887%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/bfcf6d0a-3a1f-bd50-76c1-84a40b5dd769%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/205384a2-13da-d4b7-ed6e-741fb21f24c0%28Office.15%29.aspx)|

## See also

[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[Access Object Object Members](http://msdn.microsoft.com/library/78aaacb1-c0d3-d809-088d-d543ecd71de3%28Office.15%29.aspx)
