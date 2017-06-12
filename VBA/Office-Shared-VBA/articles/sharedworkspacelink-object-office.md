---
title: SharedWorkspaceLink Object (Office)
keywords: vbaof11.chm270000
f1_keywords:
- vbaof11.chm270000
ms.prod: office
api_name:
- Office.SharedWorkspaceLink
ms.assetid: eb36dbed-fc41-08df-3cbc-affbaf5f9784
ms.date: 06/08/2017
---


# SharedWorkspaceLink Object (Office)

Represents a URL link saved in a shared document workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the  **SharedWorkspaceLink** object to manage links to additional documents and information of interest to the members who are collaborating on the documents in the shared workspace site.

Use the  **Item** ( _index_ ) property of the **SharedWorkspaceLinks** collection to return a specific **SharedWorkspaceLink** object.

Use the  **Description** property to set the link description that appears on the **Links** tab of the **Shared Workspace** pane and on the workspace Web page. Use the **URL** property to set the destination address of the link. Use the **Notes** property to supply additional information about the link.

Use the  **Save** method to upload changes to the server after you modify properties of the **SharedWorkspaceLink** object.

Use the  **CreatedBy**, **CreatedDate**, **ModifiedBy**, and **ModifiedDate** properties to return information about the history of each link.


## Example

The following example modifies the first link in the shared workspace site to point to the Microsoft Developer Network home page, then uploads the changes to the server.


```
    Dim swsLink As Office.SharedWorkspaceLink 
    Set swsLink = ActiveWorkbook.SharedWorkspace.Links(1) 
    With swsLink 
        .Description = "MSDN Home Page" 
        .URL = "http://msdn.microsoft.com/" 
        .Notes = "My favorite site for developers!" 
        .Save 
    End With 
    Set swsLink = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Delete](sharedworkspacelink-delete-method-office.md)|
|[Save](sharedworkspacelink-save-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspacelink-application-property-office.md)|
|[CreatedBy](sharedworkspacelink-createdby-property-office.md)|
|[CreatedDate](sharedworkspacelink-createddate-property-office.md)|
|[Creator](sharedworkspacelink-creator-property-office.md)|
|[Description](sharedworkspacelink-description-property-office.md)|
|[ModifiedBy](sharedworkspacelink-modifiedby-property-office.md)|
|[ModifiedDate](sharedworkspacelink-modifieddate-property-office.md)|
|[Notes](sharedworkspacelink-notes-property-office.md)|
|[Parent](sharedworkspacelink-parent-property-office.md)|
|[URL](sharedworkspacelink-url-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
