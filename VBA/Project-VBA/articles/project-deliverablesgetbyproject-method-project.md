---
title: Project.DeliverablesGetByProject Method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverablesGetByProject
ms.assetid: bbf626e8-a43e-dd06-dd2a-3d29aa1f0b6b
ms.date: 06/08/2017
---


# Project.DeliverablesGetByProject Method (Project)

Gets a list of all deliverables for the specified enterprise project in the XML member of the returned object. Project Professional only.


## Syntax

 _expression_. **DeliverablesGetByProject**( ** _ProjectGuid_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProjectGuid_|Required|**String**|GUID of the enterprise project.|

### Return Value

 **Object**


## Remarks

Using VBA to process the  **XML** member of the **DeliverablesGetByProject** result object requires complex and non-intuitive code. We recommend that you use the Office and SharePoint Development Tools in Visual Studio 2012 to create an add-in for Project when you use Project Server and SharePoint features. The easiest approach to processing XML is to use the LINQ to XML methods in the .NET Framework 4.


## Example

In the following example, the enterprise project named Simple includes a deliverable that is set for a milestone named M1. The Project site URL, which contains the Deliverables list for the Simple project, is  `http://ServerName/PWA/Simple`. The  **TestDeliverables** macro shows a message box that contains part of the XML result.

The  **projectGuid** value returned by the **GetServerProjectGuid** method includes braces around the GUID, for example, "{1b14e65c-5601-4565-acb9-3822078a17fb}". You can use a GUID value either with or without the braces.




```vb
Option Explicit 
 
Sub TestDeliverables() 
    Dim projectGuid As String 
    Dim ds As Object 
 
    projectGuid = ActiveProject.GetServerProjectGuid 
 
    ' Optional: Removing the braces on the GUID value makes no difference. 
    ' projectGuid = Mid(projectGuid, 2, 36) 
 
    Set ds = ActiveProject.DeliverablesGetByProject(projectGuid) 
 
    MsgBox ds.XML 
 
    Debug.Print ds.XML 
End Sub
```


 **Note**  To find members of a variable of type  **Object**, such as the **ds** variable, set a watch on the object, and then set a breakpoint after you assign a value to the object. Expand the variable in the **Watch** pane, and you can see the **XML** member.

The message box shows only the first 1024 characters of the total 17,295 characters of the XML result (in this example). In the following XML result, attributes are broken into separate lines. The actual XML result is all on one line, which you can see if you print the result to the  **Immediate** pane in the VBE. The example does not show the XML schema, which makes up most of the content.

The  **ows_** fields are defined in the SharePoint list. Some fields that you may want to extract include **deliverableUid**, **workspaceUri**, **linkedTaskUid** (GUID of the task in Project Server), **ows_LinkTitle** (the name of the task that has the deliverable), **ows_Created**, **ows_Modified**, **ows_Author**, **ows_CommitmentStart**, and **ows_CommitmentFinish**.




```XML
<DeliverableMasterDocument> 
 <Deliverables> 
 <Deliverable deliverableUid="6f8cb9a5-d9b8-496d-af90-1e88dc57f46a" projectUid="1b14e65c-5601-4565-acb9-3822078a17fb" 
 type="1" tpId="1" workspaceUri="http://ServerName/PWA/Simple" workspaceName="PWA/Simple" workspaceVServerUri="http://ServerName" 
 listUid="168a6e6f-6993-4315-a593-7ffa21683e57" state="1"> 
 <Client linkedTaskUid="d3eaf532-9ab9-4eb2-8f85-fd41a1b5db0c" ows_ID="1" 
 ows_ContentTypeId="0x010074416DB49FB844B99C763FA7171E7D1F00001031A192BFCA4D83CA160D2BCAB735" 
 ows_ContentType="Project Site Deliverable" ows_Title="M1" ows_Modified="2010-02-19 13:30:19" 
 ows_Created="2010-02-19 13:29:45" ows_Author="1073741823;#System Account" 
 ows_Editor="1073741823;#System Account" ows_owshiddenversion="2" ows_WorkflowVersion="1" 
 ows__UIVersion="512" ows__UIVersionString="1.0" ows_Attachments="0" ows__ModerationStatus="0" 
 ows_LinkTitleNoMenu="M1" ows_LinkTitle="M1" ows_LinkTitle2="M1" ows_SelectTitle="1" 
 ows_Order="100.000000000000" ows_GUID="{FFA3E0F9-DBB4-44B6-B09D-1C2AB7A9EF92}" 
 ows_FileRef="1;#PWA/Simple/Lists/Deliverables/1_.000" ows_FileDirRef="1;#PWA/Simple/Lists/Deliverables" 
 ows_Last_x0020_Modified="1;#2010-02-19 13:29:45" ows_Created_x0020_Date="1;#2010-02-19 13:29:45" 
 ows_FSObjType="1;#0" ows_SortBehavior="1;#0" ows_PermMask="0x7fffffffffffffff" ows_FileLeafRef="1;#1_.000" 
 ows_UniqueId="1;#{29AF34EA-EA27-48C7-80A6-83B0A95DB9BD}" ows_ProgId="1;#" 
 ows_ScopeId="1;#{73C1A12E-DBA2-4BE2-87EE-1FF5EF1494DD}" ows__EditMenuTableStart="1_.000" 
 ows__EditMenuTableStart2="1" ows__EditMenuTableEnd="1" ows_LinkFilenameNoMenu="1_.000" 
 ows_LinkFilename="1_.000" ows_LinkFilename2="1_.000" ows_ServerUrl="/PWA/Simple/Lists/Deliverables/1_.000" 
 ows_EncodedAbsUrl="http://jc2vm1/PWA/Simple/Lists/Deliverables/1_.000" ows_BaseName="1_" ows_MetaInfo="1;#" 
 ows__Level="1" ows__IsCurrentVersion="1" ows_ItemChildCount="1;#0" ows_FolderChildCount="1;#0" 
 ows_CommitmentStart="2010-02-02 00:00:00" ows_CommitmentFinish="2010-02-02 00:00:00" ows_SuppressCreateEvent="1"/> 
 </Deliverable> 
 </Deliverables> 
 <Schemas> 
 <Schema . . . 
 . . . > 
 <Fields> 
 <Field . . . /> 
 . . . 
 </Fields> 
 </Schema> 
 </Schemas> 
</DeliverableMasterDocument>
```


