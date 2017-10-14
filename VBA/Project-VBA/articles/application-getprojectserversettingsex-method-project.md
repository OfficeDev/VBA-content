---
title: Application.GetProjectServerSettingsEx Method (Project)
ms.prod: project-server
api_name:
- Project.Application.GetProjectServerSettingsEx
ms.assetid: cd630197-60e0-0ba8-e01e-114b82fe9f1e
ms.date: 06/08/2017
---


# Application.GetProjectServerSettingsEx Method (Project)

Returns global Project settings in a single XML string. You can obtain settings specific to the active project, or you can obtain settings specific to the current project manager by calling a server-side object.


## Syntax

 _expression_. **GetProjectServerSettingsEx**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **String**


## Remarks

Following is the format of the the XML string returned by  **GetProjectServerSettingsEx** method.


```XML
<ProjectServerSettingsReturn> 
   <ProjectIDInProjectServer>self explanatory</ProjectIDInProjectServer> 
   <AdminDefaultTrackingMethod>see explanation below</AdminDefaultTrackingMethod> 
   <AdminTrackingLocked>(1 or 0)</AdminTrackingLocked> 
   <ProjectManagerHasTransactions>(number of transactions)</ProjectManagerHasTransactions> 
   <ProjectManagerHasTransactionsForCurrentProject>(number of pending transactions)</ProjectManagerHasTransactionsForCurrentProject> 
   <GroupsForCurrentProjectManager> 
      <ProjectServerGroup>Name of first group that user belongs to</ProjectServerGroup> 
       ... 
      <ProjectServerGroup>Name of nth group that user belongs to</ProjectServerGroup> 
   </GroupsForCurrentProjectManager> 
</ProjectServerSettingsReturn>
```

 **Where:**

 **ProjectIDInProjectServer -** The class identifier of the active project.

 **AdminDefaultTrackingMethod -** Default tracking method for task status. You can see this on Microsoft Project Web Access by going to **Server Settings**->( **Time and task management section**)  **Task Settings and Display**-> **Tracking Method** property. It can be one of the following:


- 1 = Hours of work done per day. Resources report their hours worked on each task per day.)
    
- 2 = Percent of work complete. Resources report the percent of work they have completed, from 0 through 100 percent)
    
- 3 = Actual work done and work remaining. Resources report the actual work done and the work remaining to be done on each task.)
    


 **AdminTrackingLocked -** Whether or not managers are forced to use the tracking method specified on the server for all projects. You can see this on Project Web App by going to ** Server Settings->(Time and task management section)Task Settings and Display->Tracking Method** property. It can be one of the following:


- 0 = Managers are not forced.
    
- 1 = Managers are forced.
    


 **ProjectManagerHasTransactions -** This returns the number of status updates that the project manager has for the active project. In Microsoft Office Project 2003, users can pass in a project ID as part of the XML parameter, but in later versions of Project the project ID is ignored.

 **ProjectManagerHasTransactionsForCurrentProject -** Returns the number of status updates that the project manager has for the active project.

 **GroupsForCurrentProjectManager -** The security groups that the project manager is a member of.


