---
title: TaskDependency Object (Project)
ms.prod: project-server
api_name:
- Project.TaskDependency
ms.assetid: 05d759fb-0203-761e-10f3-65b07d233f4d
ms.date: 06/08/2017
---


# TaskDependency Object (Project)



Represents the link type and link lag information between two tasks. The  **TaskDependency** object is a member of the **[TaskDependencies](taskdependencies-object-project.md)** collection.
 **Using the TaskDependency Object**
Use  **TaskDependencies** ( _Index_ ), where _Index_ is the dependency index, to return a single **TaskDependency** object. The following example adds 1.5 days of lag to the link between the specified task and the predecessor specified in its first task dependency.
 **Using the TaskDependencies Collection**
Use the  **[TaskDependencies](http://msdn.microsoft.com/library/9c02fe5f-cb9e-a10e-bf9a-66b7600f8c64%28Office.15%29.aspx)** property to return a **TaskDependencies** collection. The following example examines each predecessor for the specified task and displays a message for each that has a priority of "High" or better.
Use the  **[Add](http://msdn.microsoft.com/library/37e67ab2-ca7b-26c2-50e7-8a933b746489%28Office.15%29.aspx)** method to add a **TaskDependency** object to the **TaskDependencies** collection. The following example links "Preliminary Research &amp; Approval" as a predecessor to "Draft Initial Business Case" in a finish-to-start relationship.

## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/02ed131a-8035-5074-e88c-f0c64e6808ad%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/86e0bda9-123b-989d-e173-4d7224fc36b9%28Office.15%29.aspx)|
|[From](http://msdn.microsoft.com/library/76127fff-e8c0-f5b4-da5b-510a5f2222fa%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/709c6af8-e383-8d41-e4d5-2e928d450905%28Office.15%29.aspx)|
|[Lag](http://msdn.microsoft.com/library/d3370ea3-5485-24d5-e363-ec4b5a0ec95b%28Office.15%29.aspx)|
|[LagType](http://msdn.microsoft.com/library/0c055a94-ea5f-1267-0b61-d3a50c6bc9b4%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/74ee0cd7-07cd-6be3-1e11-06b0eede5373%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/d6007a61-9079-7a19-93ea-94f3d6e880f1%28Office.15%29.aspx)|
|[To](http://msdn.microsoft.com/library/b2b26a7c-cbbd-c61c-a598-a04d9628fe0f%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/fb8203b5-72ab-8b10-6698-461a75fce588%28Office.15%29.aspx)|

