---
title: Subproject Object (Project)
ms.prod: project-server
api_name:
- Project.Subproject
ms.assetid: 1a3b0d18-6464-a4f2-479f-710e19faffa8
ms.date: 06/08/2017
---


# Subproject Object (Project)



Represents a subproject. The  **Subproject** object is a member of the **[Subprojects](subprojects-object-project.md)** collection.
 **Using the Subproject Object**
Use  **Subprojects** ( _Index_ ), where _Index_ is the subproject index or project summary task name, to return a single **Subproject** object. The following example prevents changes made to the specified subproject in a master project from being automatically made to the source project.
 **Using the Subprojects Collection**
Use the  **[Subprojects](http://msdn.microsoft.com/library/e4b143fb-3da7-69bd-6535-5604c2cc2dc0%28Office.15%29.aspx)** property to return a **Subprojects** collection. The following example cautions the user if any of the subprojects in the active project are not on the hard disk.

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/412c720b-a432-6e3f-96b3-f6e16c3ee48c%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/90cb228c-e757-3826-7735-5ff169477171%28Office.15%29.aspx)|
|[InsertedProjectSummary](http://msdn.microsoft.com/library/a98d0c9c-2c9d-d15e-2716-ed27ee9273c2%28Office.15%29.aspx)|
|[IsLoaded](http://msdn.microsoft.com/library/5e2e5877-1e60-9797-3fc9-ab10d8a64c1c%28Office.15%29.aspx)|
|[LinkToSource](http://msdn.microsoft.com/library/8055fc21-1de2-dbd1-c28d-2200e8bc781d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/5676f800-20ce-7607-cdec-ea7596eb1cb5%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/57bd6c44-5a2e-a2c8-c733-4c46e32be780%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/a42bc4d7-bd50-5846-76c8-27c32713bfab%28Office.15%29.aspx)|
|[SourceProject](http://msdn.microsoft.com/library/4135a5c9-eacb-12d3-b631-1d30d689f666%28Office.15%29.aspx)|

