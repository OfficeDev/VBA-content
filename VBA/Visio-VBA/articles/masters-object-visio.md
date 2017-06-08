---
title: Masters Object (Visio)
keywords: vis_sdr.chm10135
f1_keywords:
- vis_sdr.chm10135
ms.prod: visio
api_name:
- Visio.Masters
ms.assetid: 07c80948-8cee-34d2-dbc9-89ca031343df
ms.date: 06/08/2017
---


# Masters Object (Visio)

 Includes a **Master** object for each master in a document's stencil.


## Remarks

To retrieve a  **Masters** collection, use the **Masters** property of a **Document** object.

The default property of a  **Masters** collection is **Item**.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVMasters**
    

## Events



|**Name**|
|:-----|
|[BeforeMasterDelete](http://msdn.microsoft.com/library/6f950fa3-3cb6-d3ef-330d-2b38956d6ff3%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/3aed0ebc-3658-f9b9-ae63-dd1f0e3efe54%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/4641bec6-204c-1196-acb0-f9aa1e8de83d%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/ab9b85e4-1639-541c-0a06-19f1def31569%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/0abb97fc-ffd6-02ef-b9b3-bbad421c1daf%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/1ebdad8c-5073-7f6c-d811-42d3725776ad%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/bf2ed2be-276a-04d8-cd98-70929cfd31f6%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/76f8d86d-dfe9-7749-ae33-96bec632d47a%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/da0e566a-a89d-c77d-d966-73d87f5eb131%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/dbdecd35-1996-465d-afd3-a82e6bb14f7b%28Office.15%29.aspx)|
|[MasterAdded](http://msdn.microsoft.com/library/d6374a9e-1c15-73b0-086c-5f511943aeec%28Office.15%29.aspx)|
|[MasterChanged](http://msdn.microsoft.com/library/824b7d27-b687-8a35-b97c-f4cf5e269065%28Office.15%29.aspx)|
|[MasterDeleteCanceled](http://msdn.microsoft.com/library/8af99a47-397c-b4f1-99d8-06bef4f8b7f0%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/11ce64dc-a7d2-cb63-1c1b-d2d99dad5525%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/c4f30992-b598-048c-6b68-30cedcef3353%28Office.15%29.aspx)|
|[QueryCancelMasterDelete](http://msdn.microsoft.com/library/69aa351f-2e89-545d-0cf8-f650d532d3a6%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/2c9790f4-4eae-0f78-e651-d5f010b019fb%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/bda14051-5cca-ba25-1b33-14514d6f5fa6%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/51a863e6-16ff-f7f1-922f-605631486176%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/d152ee14-96e0-7cde-6a9f-2ea16d17799f%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/378f6a8f-f434-3c80-b2b2-9bde768a2f09%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/81f3c6b1-0148-aa72-716f-d24484e6710b%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/8a3c90af-47c1-440c-fb91-d16ebfabd2df%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/d4237896-734b-5308-d5db-bceef77f6b57%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/5c838330-1d66-d343-0a50-846c91496325%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/b01fb699-4c8b-2f86-c69d-70aee941c49b%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/d443f6e0-0bd9-bd55-15bf-f34e17b22ad5%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/3951e242-c7e6-7a30-bf2c-0af7c030ace1%28Office.15%29.aspx)|
|[AddEx](http://msdn.microsoft.com/library/a27b1a7c-37f4-90c9-91f1-2249611a3cf9%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/aff32258-755c-cce3-5f46-e611de6c8f5a%28Office.15%29.aspx)|
|[GetNames](http://msdn.microsoft.com/library/3cdea9a5-97da-4f59-2a93-7a1d15c29e54%28Office.15%29.aspx)|
|[GetNamesU](http://msdn.microsoft.com/library/bf797a6a-1018-eda6-41e8-c8533638a034%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/fb355d9b-7b5f-469e-3a75-f1b0fed7300b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/e7962cea-2747-82d5-50a9-73f571513247%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/baf61642-ccf8-ad9e-b131-8741f3b2c8ba%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/51130b43-b795-eb51-41c2-c7bd60f03766%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/1703269d-91bb-2a66-538c-20aecd48f879%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/20837fbb-56d0-b23c-f7de-8fd3d7a99b15%28Office.15%29.aspx)|
|[ItemFromID](http://msdn.microsoft.com/library/50cae679-5a81-ae45-6e61-8ec914f525f0%28Office.15%29.aspx)|
|[ItemU](http://msdn.microsoft.com/library/fa4e26a1-21d1-04bf-4fd8-83049cc0a5df%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/c8dc1643-1ff5-8c81-6fd0-be3c8b569443%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/87c2ab38-875a-5485-22b5-f936b84201b8%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/626b520d-ce0b-40b4-1a46-11fa9a59b0b7%28Office.15%29.aspx)|

