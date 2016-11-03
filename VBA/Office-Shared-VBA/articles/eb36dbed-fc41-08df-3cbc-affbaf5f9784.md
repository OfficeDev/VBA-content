
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
|[Delete](8db5de1e-7dc3-ebcc-1853-69b6f382d19d.md)|
|[Save](5e5f2d01-19dd-a7fb-04aa-25cacb53c02e.md)|

## Properties



|**Name**|
|:-----|
|[Application](28c2a2b5-b709-e0be-f8a5-dc6b679185f4.md)|
|[CreatedBy](a97760f8-5bed-7834-4890-21ef211cee32.md)|
|[CreatedDate](4d3a905c-4472-d0e9-ad2d-556ec34b1801.md)|
|[Creator](f6e91cf1-ceca-d5b6-d71e-26253943e429.md)|
|[Description](0f03cbdc-228d-0580-23b5-d6b4c9f4ee66.md)|
|[ModifiedBy](3070460c-c3af-ff17-19b7-25a3c6339628.md)|
|[ModifiedDate](0ad877d1-a1dd-558d-eee0-9502f8242b6b.md)|
|[Notes](5bb05b61-2746-f276-5159-ee8f28a30c66.md)|
|[Parent](a6470d25-9f45-c90d-4feb-ff823f969883.md)|
|[URL](92104c43-43b8-5f59-e0c0-91313d8f5e35.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)