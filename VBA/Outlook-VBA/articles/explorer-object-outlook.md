---
title: Explorer Object (Outlook)
keywords: vbaol11.chm2985
f1_keywords:
- vbaol11.chm2985
ms.prod: outlook
api_name:
- Outlook.Explorer
ms.assetid: 026591e5-049f-503a-4166-34e6dbc225fb
ms.date: 06/08/2017
---


# Explorer Object (Outlook)

Represents the window in which the contents of a folder are displayed.


## Remarks




- Use the  **[Item](http://msdn.microsoft.com/library/b854ab0e-e966-4de8-7ccf-db4723812212%28Office.15%29.aspx)** method of the **[Explorers](http://msdn.microsoft.com/library/8398532a-1fad-7390-6778-109ac5e6c67c%28Office.15%29.aspx)** object to return the object representing a specific explorer.
    
- Use the  **[ActiveExplorer](http://msdn.microsoft.com/library/f6dd27c0-4319-c7fc-191f-8b3b2ea319d3%28Office.15%29.aspx)** method to return the object representing the currently active explorer (if there is one).
    
- Use the  **[GetExplorer](http://msdn.microsoft.com/library/f60bf373-802e-cb93-2152-bc6c8945edb1%28Office.15%29.aspx)** method to return the **Explorer** object associated with a folder.
    
- Use the  **[Display](http://msdn.microsoft.com/library/cde389e0-5ec9-8261-5ec0-9a5ba4f8776d%28Office.15%29.aspx)** method of a **[Folder](folder-object-outlook.md)** object to display a folder in its associated explorer.
    

## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/8543d347-baf5-cdc9-2366-11c9917e035e%28Office.15%29.aspx)|
|[AttachmentSelectionChange](http://msdn.microsoft.com/library/9694482b-657c-82d5-9ad6-c1df644795b2%28Office.15%29.aspx)|
|[BeforeFolderSwitch](http://msdn.microsoft.com/library/ae65c073-6b4a-ac81-c4ae-691118b19df0%28Office.15%29.aspx)|
|[BeforeItemCopy](http://msdn.microsoft.com/library/05ae7be8-5528-5560-f8ce-73f0afbf4cde%28Office.15%29.aspx)|
|[BeforeItemCut](http://msdn.microsoft.com/library/82861e5e-e990-aed9-4134-db9cbe63d47c%28Office.15%29.aspx)|
|[BeforeItemPaste](http://msdn.microsoft.com/library/a6d43429-5309-4b07-7b0b-68cddd2d7e59%28Office.15%29.aspx)|
|[BeforeMaximize](http://msdn.microsoft.com/library/4d55aa87-44c6-4660-c2bf-579d3b9dc376%28Office.15%29.aspx)|
|[BeforeMinimize](http://msdn.microsoft.com/library/999b2bc3-99de-6dc8-81a2-dd25c8bc71c6%28Office.15%29.aspx)|
|[BeforeMove](http://msdn.microsoft.com/library/bce617d3-3bf8-2a59-ab0a-4ef1e7759c75%28Office.15%29.aspx)|
|[BeforeSize](http://msdn.microsoft.com/library/2df91a98-89e2-82af-acfc-49f8e9f40952%28Office.15%29.aspx)|
|[BeforeViewSwitch](http://msdn.microsoft.com/library/5b7ac070-ba4d-6fa8-94e5-20370efe7343%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/20586ee0-35b5-02f9-327b-8431f6083cca%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/7bf07653-3e12-670b-c293-1d51cf30e564%28Office.15%29.aspx)|
|[FolderSwitch](http://msdn.microsoft.com/library/5dfa1fa3-c381-8e19-0528-d70a6fd63187%28Office.15%29.aspx)|
|[InlineResponse](http://msdn.microsoft.com/library/5dbaddbd-e6cd-4776-b417-c67f51b12812%28Office.15%29.aspx)|
|[InlineResponseClose](http://msdn.microsoft.com/library/ff3f3286-995a-409c-aca5-706290e26252%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/ef0d976f-b9f6-2080-7657-e48d1c64ccb1%28Office.15%29.aspx)|
|[ViewSwitch](http://msdn.microsoft.com/library/ab981f42-d429-ccd7-a25c-142e52683020%28Office.15%29.aspx)|
|[DisplayModeChange](http://msdn.microsoft.com/library/cee77aad-8905-efed-466e-c2e88cfeeaa2%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/53f33d64-7a33-6772-4abc-fe328d3abb57%28Office.15%29.aspx)|
|[AddToSelection](http://msdn.microsoft.com/library/b85ad121-9e26-0782-3c5e-7651499f8e66%28Office.15%29.aspx)|
|[ClearSearch](http://msdn.microsoft.com/library/644b6012-0b87-b4cb-6104-6f05b5c4dcc5%28Office.15%29.aspx)|
|[ClearSelection](http://msdn.microsoft.com/library/2809b5fb-961e-fb2a-a74d-fffa4484c838%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/df5ecd62-066a-0b46-3a5c-e7d955677f4a%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/3d93be5a-90af-af60-c16a-ec15d87f4d97%28Office.15%29.aspx)|
|[IsItemSelectableInView](http://msdn.microsoft.com/library/a2ec8bbb-0f24-6db6-05a8-1b8375b71da7%28Office.15%29.aspx)|
|[IsPaneVisible](http://msdn.microsoft.com/library/d547978a-f6b4-06ea-2358-8b6a81230240%28Office.15%29.aspx)|
|[RemoveFromSelection](http://msdn.microsoft.com/library/f31bc78f-500e-2f73-ea14-8d5f19cd44e9%28Office.15%29.aspx)|
|[Search](http://msdn.microsoft.com/library/d4dc7ae5-c24f-90df-f52e-e0b73293e25d%28Office.15%29.aspx)|
|[SelectAllItems](http://msdn.microsoft.com/library/05b3169a-5f27-2169-5ac5-1d64951d6430%28Office.15%29.aspx)|
|[ShowPane](http://msdn.microsoft.com/library/3d2c9dd5-b660-e160-36db-73c23f95a7a2%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccountSelector](http://msdn.microsoft.com/library/5d383684-a88e-8266-522b-7762895e69d3%28Office.15%29.aspx)|
|[ActiveInlineResponse](http://msdn.microsoft.com/library/fc38314d-7cff-44f4-9151-6129f918a721%28Office.15%29.aspx)|
|[ActiveInlineResponseWordEditor](http://msdn.microsoft.com/library/b9058694-ab8f-4962-ab7d-afac1704dd29%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/d3318c7b-55c4-7797-7abf-c2c71911fb01%28Office.15%29.aspx)|
|[AttachmentSelection](http://msdn.microsoft.com/library/d516b972-5eb0-7a76-d4b6-000e26d523aa%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/69f20794-7b31-4999-3c2f-525f1a15f7f6%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/12873732-cb5f-e6ca-1328-05cf908038e5%28Office.15%29.aspx)|
|[CurrentFolder](http://msdn.microsoft.com/library/75e7f120-28df-0c3b-ec05-bd880621141b%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/177e6387-9ccb-cb71-bbe5-332c25485848%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/bce6fc29-c52b-13da-d68b-4b45b694e880%28Office.15%29.aspx)|
|[HTMLDocument](http://msdn.microsoft.com/library/dd9ff575-37f5-1b64-5ebf-f17998586d28%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/83691416-276b-a77f-4a20-9fc2443571e0%28Office.15%29.aspx)|
|[NavigationPane](http://msdn.microsoft.com/library/9ff92a76-d1cd-e338-2f45-e3e5c79c136e%28Office.15%29.aspx)|
|[Panes](http://msdn.microsoft.com/library/b7ec51bd-c8e0-f31e-1f15-42a7514cb433%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/32fc387d-a3f2-05b4-ffaf-f93c50f51406%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/11002043-9dab-a5ad-b36e-52ddb04c1859%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/47752d87-6ef5-4838-4c08-0325c0b613f7%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/f3afa2a5-e532-072d-1be0-4600c4848031%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/7e5caaf7-c572-d74a-1019-e9fc2cf78d84%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/787b6339-eb92-3ab6-df9f-82f6122facc5%28Office.15%29.aspx)|
|[DisplayMode](http://msdn.microsoft.com/library/8e6bcc0d-5a37-2c8f-d059-28706b638dee%28Office.15%29.aspx)|
|[PreviewPane](http://msdn.microsoft.com/library/5f3edb49-b9f6-db03-8f83-3fe27f0aaf08%28Office.15%29.aspx)|

## See also


#### Other resources


[Explorer Object Members](http://msdn.microsoft.com/library/4412c507-4dcd-6005-b9c8-11824624250d%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
