
# Masters Members (Visio)
 Includes a **Master** object for each master in a document's stencil.

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Events](#sectionSection0)
 [Methods](#sectionSection1)
 [Properties](#sectionSection2)



## Events
<a name="sectionSection0"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [BeforeMasterDelete](6f950fa3-3cb6-d3ef-330d-2b38956d6ff3.md)|Occurs before a master is deleted from a document.|
| [BeforeSelectionDelete](3aed0ebc-3658-f9b9-ae63-dd1f0e3efe54.md)|Occurs before selected objects are deleted.|
| [BeforeShapeDelete](4641bec6-204c-1196-acb0-f9aa1e8de83d.md)|Occurs before a shape is deleted.|
| [BeforeShapeTextEdit](ab9b85e4-1639-541c-0a06-19f1def31569.md)|Occurs before a shape is opened for text editing in the user interface.|
| [CellChanged](0abb97fc-ffd6-02ef-b9b3-bbad421c1daf.md)|Occurs after the value changes in a cell in a document.|
| [ConnectionsAdded](1ebdad8c-5073-7f6c-d811-42d3725776ad.md)|Occurs after connections have been established between shapes.|
| [ConnectionsDeleted](bf2ed2be-276a-04d8-cd98-70929cfd31f6.md)|Occurs after connections between shapes have been removed.|
| [ConvertToGroupCanceled](76f8d86d-dfe9-7749-ae33-96bec632d47a.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
| [FormulaChanged](da0e566a-a89d-c77d-d966-73d87f5eb131.md)|Occurs after a formula changes in a cell in the object that receives the event.|
| [GroupCanceled](dbdecd35-1996-465d-afd3-a82e6bb14f7b.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
| [MasterAdded](d6374a9e-1c15-73b0-086c-5f511943aeec.md)|Occurs after a new master is added to a document.|
| [MasterChanged](824b7d27-b687-8a35-b97c-f4cf5e269065.md)|Occurs after properties of a master are changed and propagated to its instances.|
| [MasterDeleteCanceled](8af99a47-397c-b4f1-99d8-06bef4f8b7f0.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.|
| [QueryCancelConvertToGroup](11ce64dc-a7d2-cb63-1c1b-d2d99dad5525.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelGroup](c4f30992-b598-048c-6b68-30cedcef3353.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelMasterDelete](69aa351f-2e89-545d-0cf8-f650d532d3a6.md)|Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelSelectionDelete](2c9790f4-4eae-0f78-e651-d5f010b019fb.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelUngroup](bda14051-5cca-ba25-1b33-14514d6f5fa6.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [SelectionAdded](51a863e6-16ff-f7f1-922f-605631486176.md)|Occurs after one or more shapes are added to a document.|
| [SelectionDeleteCanceled](d152ee14-96e0-7cde-6a9f-2ea16d17799f.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
| [ShapeAdded](378f6a8f-f434-3c80-b2b2-9bde768a2f09.md)|Occurs after one or more shapes are added to a document.|
| [ShapeChanged](81f3c6b1-0148-aa72-716f-d24484e6710b.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
| [ShapeDataGraphicChanged](8a3c90af-47c1-440c-fb91-d16ebfabd2df.md)|Occurs after a data graphic is applied to or deleted from a shape.|
| [ShapeExitedTextEdit](d4237896-734b-5308-d5db-bceef77f6b57.md)|Occurs after a shape is no longer open for interactive text editing.|
| [ShapeParentChanged](5c838330-1d66-d343-0a50-846c91496325.md)|Occurs after shapes are grouped or a group is ungrouped.|
| [TextChanged](b01fb699-4c8b-2f86-c69d-70aee941c49b.md)|Occurs after the text of a shape is changed in a document.|
| [UngroupCanceled](d443f6e0-0bd9-bd55-15bf-f34e17b22ad5.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods
<a name="sectionSection1"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [Add](3951e242-c7e6-7a30-bf2c-0af7c030ace1.md)|Adds a new object to a collection.|
| [AddEx](a27b1a7c-37f4-90c9-91f1-2249611a3cf9.md)|Adds a new  **Master** object of the specified type to the **Masters** collection of a Microsoft Visio document.|
| [Drop](aff32258-755c-cce3-5f46-e611de6c8f5a.md)|Creates a new **Master** object by dropping an object onto a receiving object such as a stencil or document, or the **Masters** or **MasterShortcuts** collection.|
| [GetNames](3cdea9a5-97da-4f59-2a93-7a1d15c29e54.md)|Returns the names of all items in a collection.|
| [GetNamesU](bf797a6a-1018-eda6-41e8-c8533638a034.md)|Returns the universal names of all items in a collection.|
| [Paste](fb355d9b-7b5f-469e-3a75-f1b0fed7300b.md)|Pastes the contents of the Clipboard into an object.|

## Properties
<a name="sectionSection2"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [Application](e7962cea-2747-82d5-50a9-73f571513247.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
| [Count](baf61642-ccf8-ad9e-b131-8741f3b2c8ba.md)|Returns the number of objects in a collection. Read-only.|
| [Document](51130b43-b795-eb51-41c2-c7bd60f03766.md)|Gets the  **Document** object that is associated with an object. Read-only.|
| [EventList](1703269d-91bb-2a66-538c-20aecd48f879.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
| [Item](20837fbb-56d0-b23c-f7de-8fd3d7a99b15.md)|Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.|
| [ItemFromID](50cae679-5a81-ae45-6e61-8ec914f525f0.md)|Returns an item of a collection using the ID of the item. Read-only.|
| [ItemU](fa4e26a1-21d1-04bf-4fd8-83049cc0a5df.md)|Returns an object from a collection. Read-only.|
| [ObjectType](c8dc1643-1ff5-8c81-6fd0-be3c8b569443.md)|Returns an object's type. Read-only.|
| [PersistsEvents](87c2ab38-875a-5485-22b5-f936b84201b8.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
| [Stat](626b520d-ce0b-40b4-1a46-11fa9a59b0b7.md)|Returns status information for an object. Read-only.|
