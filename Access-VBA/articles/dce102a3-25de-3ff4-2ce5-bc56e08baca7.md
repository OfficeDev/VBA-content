
# Workbook Members (Excel)
Represents a Microsoft Excel workbook.

 **Last modified:** July 28, 2015

 **In this article**
 [Events](#sectionSection0)
 [Methods](#sectionSection1)
 [Properties](#sectionSection2)


## Events
<a name="sectionSection0"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [Activate](74bb6d8c-aec8-7bb6-5c30-9a20f9a7afe8.md)|Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.|
| [AddinInstall](671117b2-590e-9d6f-29ae-5f0bf30d4e99.md)|Occurs when the workbook is installed as an add-in|
| [AddinUninstall](e35ba67b-3e04-d950-2f8b-141e478ddb67.md)|Occurs when the workbook is uninstalled as an add-in.|
| [AfterSave](97fee36a-f77c-29ab-de1d-b6069b2d74d8.md)|Occurs after the workbook is saved.|
| [AfterXmlExport](fe1e0a53-9f4e-ac88-58f7-fe420e57cabd.md)|Occurs after Microsoft Excel saves or exports XML data from the specified workbook. |
| [AfterXmlImport](b43adf53-6b67-6127-e69d-6ea05f68b7f6.md)|Occurs after an existing XML data connection is refreshed or after new XML data is imported into the specified Microsoft Excel workbook.|
| [BeforeClose](1c440637-8289-c6dd-24e0-1b2764fd1694.md)|Occurs before the workbook closes. If the workbook has been changed, this event occurs before the user is asked to save changes.|
| [BeforePrint](2c97cb32-2bb3-2848-b5ed-32d9129af080.md)|Occurs before the workbook (or anything in it) is printed.|
| [BeforeSave](dfa3e20f-1fb2-f84f-4b92-a98f22b6e637.md)|Occurs before the workbook is saved.|
| [BeforeXmlExport](ee2af5de-e52f-9434-aa7c-5dc9bb102d1b.md)|Occurs before Microsoft Excel saves or exports XML data from the specified workbook.|
| [BeforeXmlImport](a0a589c6-15f9-5599-c0b6-c6f881816ad6.md)|Occurs before an existing XML data connection is refreshed or before new XML data is imported into a Microsoft Excel workbook.|
| [Deactivate](6bd5411c-ac43-95cf-6755-49780ac765e9.md)|Occurs when the chart, worksheet, or workbook is deactivated.|
| [ModelChange](efe01088-273b-f9d8-ea3e-2ea1725ba7b2.md)|Occurs after the Excel data model is changed. |
| [NewChart](76e7f325-9244-fd8c-b38d-063f0193a5e9.md)|Occurs when a new chart is created in the workbook.|
| [NewSheet](5abb254d-a2c3-7dac-e79f-0de74a081ecd.md)|Occurs when a new sheet is created in the workbook.|
| [Open](313adc5e-0319-4ca4-cf5d-791b7184dacf.md)|Occurs when the workbook is opened.|
| [PivotTableCloseConnection](e267ab5b-382e-b270-18c8-f643e03e4604.md)|Occurs after a PivotTable report closes the connection to its data source.|
| [PivotTableOpenConnection](b6ce12f7-7bc6-bfcc-33f4-2e8ea6e53bae.md)|Occurs after a PivotTable report opens the connection to its data source.|
| [RowsetComplete](05bdddba-6716-4bba-01b6-863f27623821.md)|The event is raised when the user either drills through the recordset or invokes the rowset action on an OLAP PivotTable.|
| [SheetActivate](2a7c05c3-5b66-8012-5ac5-981dcfc7f947.md)|Occurs when any sheet is activated.|
| [SheetBeforeDelete](42406738-0fcd-4ef7-9bd6-abcc05f5e922.md)||
| [SheetBeforeDoubleClick](69d21025-78ef-deab-39be-b7a092d611f5.md)|Occurs when any worksheet is double-clicked, before the default double-click action.|
| [SheetBeforeRightClick](d84dd9fd-85d3-009e-281b-cfc0d2874859.md)|Occurs when any worksheet is right-clicked, before the default right-click action.|
| [SheetCalculate](0610bfa5-15dc-a57f-f362-cf897bd54b91.md)|Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.|
| [SheetChange](37e727d8-255c-ac23-45d8-13a8e7639991.md)|Occurs when cells in any worksheet are changed by the user or by an external link.|
| [SheetDeactivate](befde22b-69ce-c34f-2b9e-da5e026972e3.md)|Occurs when any sheet is deactivated.|
| [SheetFollowHyperlink](be29df8c-4e8e-f719-ae1d-f91a11b89491.md)|Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the  ** [FollowHyperlink](c63eec19-008e-bfb5-1357-3d02426c1bab.md)**event.|
| [SheetLensGalleryRenderComplete](8ac48e9f-7a15-c674-6d96-e9c1466473bc.md)|Occurs when a callout gallery's icons (dynamic &amp; static) have completed rendering for a worksheet.|
| [SheetPivotTableAfterValueChange](8460f5f1-d415-7aac-6a3d-fa0944036e9c.md)|Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).|
| [SheetPivotTableBeforeAllocateChanges](2f767b5b-27fb-33de-c91d-76bbc52ea171.md)|Occurs before changes are applied to a PivotTable.|
| [SheetPivotTableBeforeCommitChanges](7e189a4f-a349-f862-375a-fa66311629cc.md)|Occurs before changes are committed against the OLAP data source for a PivotTable.|
| [SheetPivotTableBeforeDiscardChanges](e8f1ae21-c9ed-6f4d-a85c-d6768060a66f.md)|Occurs before changes to a PivotTable are discarded.|
| [SheetPivotTableChangeSync](c280b935-3dbf-0666-b727-64d6b4ac7ebd.md)|Occurs after changes to a PivotTable.|
| [SheetPivotTableUpdate](0b37939a-28dd-ef8b-ea5e-fc3768f8979a.md)|Occurs after the sheet of the PivotTable report has been updated.|
| [SheetSelectionChange](a3829af1-2917-9526-1d64-91eeb6c198ce.md)|Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).|
| [SheetTableUpdate](609d331e-45b9-885b-a395-d80ccf4c19a5.md)|Occurs after the sheet table has been updated.|
| [Sync](ce8b77e1-a316-c0e3-f0f8-ce4ac22ec430.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
| [WindowActivate](e99d955c-1975-44c3-05b3-3aa6e851083c.md)|Occurs when any workbook window is activated.|
| [WindowDeactivate](d84f0819-00df-585f-ea31-e4ab5a72950e.md)|Occurs when any workbook window is deactivated.|
| [WindowResize](6e473482-fe16-03a2-7a27-b0cd9535c3e6.md)|Occurs when any workbook window is resized.|

## Methods
<a name="sectionSection1"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [AcceptAllChanges](8d8572a9-1231-c8ef-0707-72b8b5109066.md)|Accepts all changes in the specified shared workbook.|
| [Activate](628e06b3-ca3f-28cb-e0fd-e696842f69f5.md)|Activates the first window associated with the workbook.|
| [AddToFavorites](14e1cd5a-41be-fc9a-61fa-df87698110e8.md)|Adds a shortcut to the workbook or hyperlink to the Favorites folder.|
| [ApplyTheme](11580293-22da-9154-20a0-6435b8870ac9.md)|Applies the specified theme to the current workbook.|
| [BreakLink](1e9d70c1-908e-92eb-26b8-d6ac753cc9c2.md)|Converts formulas linked to other Microsoft Excel sources or OLE sources to values.|
| [CanCheckIn](17f7cbdd-0ce0-8e3a-46f3-cb6dafaaa40a.md)| **True** if Microsoft Excel can check in a specified workbook to a server. Read/write **Boolean**.|
| [ChangeFileAccess](07f9cfc3-eece-efc1-6c03-38782ad7bcc2.md)|Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.|
| [ChangeLink](9b2c0b82-73ff-3bdb-63df-82c0708cb703.md)|Changes a link from one document to another.|
| [CheckIn](f9750086-aaa6-3b04-6b51-ebcadf6b1911.md)|Returns a workbook from a local computer to a server, and sets the local workbook to read-only so that it cannot be edited locally. Calling this method will also close the workbook.|
| [CheckInWithVersion](3b37cea5-8795-bcbb-9c4b-d30b2b9a095e.md)|Saves a workbook to a server from a local computer, and sets the local workbook to read-only so that it cannot be edited locally.|
| [Close](c0376cab-a2db-c606-67bf-0a4921b81e03.md)|Closes the object.|
| [DeleteNumberFormat](d56c2e4c-5de2-fecf-6a1f-a9fdc79943cb.md)|Deletes a custom number format from the workbook.|
| [EnableConnections](521ebb4c-56c6-3e21-39af-4a46934790e1.md)|The  **EnableConnections** method allows developers to programmatically enable data connections within the workbook for the user.|
| [EndReview](cd4a445b-4731-43ba-e46a-f80f19ea5a17.md)|Terminates a review of a file that has been sent for review using the  ** [SendForReview](3834f5b3-6d24-1bb9-27b5-052aa2e725e3.md)**method.|
| [ExclusiveAccess](9b92ec4f-e256-7e01-6cd7-759a0d022813.md)|Assigns the current user exclusive access to the workbook that's open as a shared list.|
| [ExportAsFixedFormat](4d72247c-bab9-3475-4792-8899c959393c.md)|The  **ExportAsFixedFormat** method is used to publish a workbook to either the PDF or XPS format.|
| [FollowHyperlink](d070ecc9-fbb6-c146-f250-5c99b09063ec.md)|Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.|
| [ForwardMailer](956b1746-26f2-5968-0ef7-fa3da2be974c.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| [GetWorkflowTasks](8a5ff9e0-b23a-930c-bb65-a1daa10cd946.md)|Returns the collection of  ** [WorkflowTask](9d17947e-f12a-2f97-7888-8d5ec9f85011.md)** objects for the specified workbook.|
| [GetWorkflowTemplates](adff72bb-39ab-69ed-8a9b-defe75a5fede.md)|Returns the collection of  ** [WorkflowTemplate](965d0474-dd51-9b0e-b34c-a11f921ff410.md)** objects for the specified workbook.|
| [HighlightChangesOptions](ac69ee3e-c5ea-5ac0-418a-0b94d56a8777.md)|Controls how changes are shown in a shared workbook.|
| [LinkInfo](016295a3-72c1-95b3-c259-8f286b12b73c.md)|Returns the link date and update status.|
| [LinkSources](6466bea0-5af8-7af0-e9d7-7595133073ae.md)|Returns an array of links in the workbook. The names in the array are the names of the linked documents, editions, or DDE or OLE servers. Returns  **Empty** if there are no links.|
| [LockServerFile](be0ac600-320e-0959-bc26-5f3f4a910f5e.md)|Locks the workbook on the server to prevent modification.|
| [MergeWorkbook](393790c6-3c19-7149-a999-b8712e7a6855.md)|Merges changes from one workbook into an open workbook.|
| [NewWindow](ba568cee-c1cb-6e6a-8935-2cc8ce3a8400.md)|Creates a new window or a copy of the specified window.|
| [OpenLinks](cae33bab-892e-0861-e4ec-8a334097e0d1.md)|Opens the supporting documents for a link or links.|
| [PivotCaches](0a2e7f10-c123-5c98-fb71-56868b9f8bde.md)|Returns a  ** [PivotCaches](cfd979b9-d52f-f34b-4b66-4fb17efcdc92.md)**collection that represents all the PivotTable caches in the specified workbook. Read-only.|
| [Post](62ecf3bc-c551-8f06-64cc-a6c141bdf172.md)|Posts the specified workbook to a public folder. This method works only with a Microsoft Exchange client connected to a Microsoft Exchange server.|
| [PrintOut](3a4e7037-fcde-5a57-4a80-45f2a0994370.md)|Prints the object.|
| [PrintPreview](044afc4c-74d6-3ea6-1811-2c7d9cdc5b1a.md)|Shows a preview of the object as it would look when printed.|
| [Protect](0e270b93-7b0b-cc68-c7c0-4002024f4292.md)|Protects a workbook so that it cannot be modified.|
| [ProtectSharing](26660bc6-136a-ffc8-987e-c96db9c08231.md)|Saves the workbook and protects it for sharing.|
| [PurgeChangeHistoryNow](7ea42af1-051b-400d-cb87-0736c49d74fb.md)|Removes entries from the change log for the specified workbook.|
| [RefreshAll](c1a956dc-263c-5c24-3b51-fc4af22dcd33.md)|Refreshes all external data ranges and PivotTable reports in the specified workbook.|
| [RejectAllChanges](a53670da-370c-9ac8-75b8-008625495c2b.md)|Rejects all changes in the specified shared workbook.|
| [ReloadAs](ce6a9d1a-7945-3dca-ff2d-a42289c2ccf9.md)|Reloads a workbook based on an HTML document, using the specified document encoding.|
| [RemoveDocumentInformation](e668d976-108b-c627-6118-dd3384c1315c.md)|Removes all information of the specified type from the workbook.|
| [RemoveUser](f0a978a0-7bcf-3af4-a01a-831c6c854989.md)|Disconnects the specified user from the shared workbook.|
| [Reply](557bb3a4-c817-e942-10cf-ba252b0db498.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| [ReplyAll](c378da35-1778-44db-5c58-8d6992ca0c93.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| [ReplyWithChanges](60424d69-0062-aa5e-ea8f-4fb07086167a.md)|Sends an e-mail message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.|
| [ResetColors](1b56a4e9-0645-fa1e-55cc-09069c6a0ff1.md)|Resets the color palette to the default colors.|
| [RunAutoMacros](85dfdadf-75e6-437d-fb7a-e17681a69b35.md)|Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook. This method is included for backward compatibility. For new Visual Basic code, you should use the Open, Close, Activate and Deactivate events instead of these macros.|
| [Save](466d891d-fb4c-fb0a-474b-dedb3c4ea7a7.md)|Saves changes to the specified workbook.|
| [SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)|Saves changes to the workbook in a different file.|
| [SaveAsXMLData](7c4c1be3-d3a5-6e90-7750-9f371f008541.md)|Exports the data that has been mapped to the specified XML schema map to an XML data file.|
| [SaveCopyAs](84f58488-6a2b-7fef-1472-e1b9771a60b0.md)|Saves a copy of the workbook to a file but doesn't modify the open workbook in memory.|
| [SendFaxOverInternet](e7d91ac4-90d2-7555-af96-dc28736da769.md)|Sends a worksheet as a fax to the specfied recipients.|
| [SendForReview](3834f5b3-6d24-1bb9-27b5-052aa2e725e3.md)|Sends a workbook in an e-mail message for review to the specified recipients.|
| [SendMail](581d197c-0748-2225-2986-64aa368aab39.md)|Sends the workbook by using the installed mail system.|
| [SendMailer](e44955e1-e250-7279-19e5-e13db80ceddc.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| [SetLinkOnData](b500a579-6e4c-5712-05cf-27c6393b3bcd.md)|Sets the name of a procedure that runs whenever a DDE link is updated.|
| [SetPasswordEncryptionOptions](3b6c9bfe-4cfb-1dde-fd57-07dd474df7db.md)|Sets the options for encrypting workbooks using passwords.|
| [ToggleFormsDesign](3a6352e3-26b9-713e-ed93-a5890b37bc0a.md)|The  **ToggleFormsDesign** method is used to toggle Excel into Design Mode when using forms controls.|
| [Unprotect](39387902-a8a4-7bf2-44d7-c5bde6725778.md)|Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.|
| [UnprotectSharing](edce1744-0906-4b4e-8b98-5d1125047bff.md)|Turns off protection for sharing and saves the workbook.|
| [UpdateFromFile](f5148b60-9b25-8a12-5cf3-40103dcff2a3.md)|Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory. If the disk copy hasn't changed since the workbook was loaded, the in-memory copy of the workbook isn't reloaded.|
| [UpdateLink](2aef72cc-a820-3e91-1f46-50c739faf2bb.md)|Updates a Microsoft Excel, DDE, or OLE link (or links).|
| [WebPagePreview](2c88f15e-5cd3-82da-f779-55b63959a2b0.md)|Displays a preview of the specified workbook as it would look if saved as a Web page.|
| [XmlImport](97964c82-1fbe-7060-0a90-23c190e0b398.md)|Imports an XML data file into the current workbook.|
| [XmlImportXml](b0edbe49-f578-ead0-8371-0196f5c515d4.md)|Imports an XML data stream that has been previously loaded into memory. Excel uses the first qualifying map found or if the destination range is specified, Excel will automatically list the data.|

## Properties
<a name="sectionSection2"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [AccuracyVersion](bc81782c-662c-87ec-8381-d06e77674d0c.md)|Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Read/write|
| [ActiveChart](81e18252-b1fe-2487-535e-6e24c80bef24.md)|Returns a  ** [Chart](179c32ce-49bd-6f36-ea12-89fb5443f3ea.md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing**.|
| [ActiveSheet](fb5578c3-64a7-edb7-4004-e608739d4c1e.md)|Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns  **Nothing** if no sheet is active.|
| [ActiveSlicer](d3858353-0be1-338c-e43f-1e5ffb7f37ba.md)|Returns an object that represents the active slicer in the active workbook or in the specified workbook. Returns  **Nothing** if no slicer is active. Read-only.|
| [Application](91b30f9d-48e5-e033-8daf-416d1c0e547d.md)|When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
| [AutoUpdateFrequency](dfded5c8-94d6-8a0f-24c1-248bd502850b.md)|Returns or sets the number of minutes between automatic updates to the shared workbook. Read/write  **Long**.|
| [AutoUpdateSaveChanges](06f9951d-a17a-bf88-4f6e-65835eb112f8.md)| **True** if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. **False** if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is **True**. Read/write  **Boolean**.|
| [BuiltinDocumentProperties](3efffd7d-0681-ecbc-000a-b71eceb3f92a.md)|Returns a  ** [DocumentProperties](90d42786-7d9a-b604-dbdf-88db41cbe69b.md)** collection that represents all the built-in document properties for the specified workbook. Read-only.|
| [CalculationVersion](09633164-998f-9fa7-f257-da109c369cd7.md)|Returns the information about the version of Excel that the workbook was last fully recalculated by. Read-only  **Long**.|
| [CaseSensitive](6053b576-9ede-f9d8-e2bf-c012653b60a2.md)| **True** if the workbook distinguishes between upper and lower case when comparing content. Read-only **Boolean**|
| [ChangeHistoryDuration](5ebc3cc5-dffa-60cf-08cb-b2f84424c4b4.md)|Returns or sets the number of days shown in the shared workbook's change history. Read/write  **Long**.|
| [ChartDataPointTrack](0aa2b1c1-0bba-f514-6158-00cdb4a5747e.md)| **True** will cause all charts in the current document to track the actual data point to which it's attached. **False** will revert back to tracking the index of the data point. **Boolean** Read/Write|
| [Charts](582d9a78-d86f-ab69-0c22-85f8a59412d9.md)|Returns a  ** [Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)** collection that represents all the chart sheets in the specified workbook.|
| [CheckCompatibility](9379c010-6756-b7ea-b4ad-5c8a4b900124.md)|Controls whether or not the compatibility checker is run automatically when the workbook is saved. Read/write  **Boolean**.|
| [CodeName](236e97b8-2bb9-c3a9-b4da-b1c327acde95.md)|Returns the code name for the object. Read-only  **String**.|
| [Colors](60fc038b-980b-c1bc-6d1c-69d9d31a11ba.md)|Returns or sets colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value. Read/write  **Variant**.|
| [CommandBars](8d93b8cd-c4e3-b216-eda0-da4c6e573c40.md)|Returns a  ** [CommandBars](0e312e21-14ee-5055-d604-b66e61c53b47.md)** object that represents the Microsoft Excel command bars. Read-only.|
| [ConflictResolution](5142c848-0731-14d9-5913-bbaa67bf308f.md)|Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated. Read/write  ** [XlSaveConflictResolution](1cdccb5a-c356-4572-9429-49850338b31b.md)**.|
| [Connections](9c4f4ba7-dd4b-0bc2-65b7-16455014097f.md)|The  **Connections** property establishes a connection between the workbook and an ODBC or an OLEDB data source and refreshes the data without prompting the user. Read-only.|
| [ConnectionsDisabled](afd53cc5-12d8-4b22-3186-1359c14f662e.md)|Disables the external connections or links in the workbook. Read-only|
| [Container](7ad370bc-9901-3b8b-12e6-1ee57f0300e0.md)|Returns the object that represents the container application for the specified OLE object. Read-only  **Object**.|
| [ContentTypeProperties](a2919232-3fa2-7f37-00c2-48eb3edb10fd.md)|Returns a  ** [MetaProperties](957a6e06-3348-b180-3655-06ffbfb69e12.md)** collection that describes the metadata stored in the workbook. Read-only.|
| [CreateBackup](33f05bf8-00ef-81f4-c083-30326f019cd4.md)| **True** if a backup file is created when this file is saved. Read-only **Boolean**.|
| [Creator](e03bdff2-7a93-f882-31a1-1ba8dd3c1764.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
| [CustomDocumentProperties](8470adbb-5b10-96ba-71f7-c667c33b6707.md)|Returns or sets a  ** [DocumentProperties](90d42786-7d9a-b604-dbdf-88db41cbe69b.md)** collection that represents all the custom document properties for the specified workbook.|
| [CustomViews](286f6d5a-fb91-a339-8e74-9014ab7f4835.md)|Returns a  ** [CustomViews](f970bdf7-371b-ba41-89a3-bef2c6907f1a.md)**collection that represents all the custom views for the workbook.|
| [CustomXMLParts](bd31f001-0e5d-691b-e69e-4cb91a6dbb0e.md)|Returns a  ** [CustomXMLParts](98c1c58e-a08d-6304-8626-1e6705917da3.md)** collection that represents the custom XML in the XML data store. Read-only.|
| [Date1904](0556311c-4e45-aea3-e922-24a5830b19d4.md)| **True** if the workbook uses the 1904 date system. Read/write **Boolean**.|
| [DefaultPivotTableStyle](8e2ca78a-4eb1-1b1e-c947-8a724f6d690a.md)|Specifies the table style from the  **TableStyles** collection that is used as the default style for PivotTables. Read/write.|
| [DefaultSlicerStyle](0f193fb8-b766-9093-9db8-8b028da108b4.md)|Specifies the style from the  ** [TableStyles](952da370-51cb-b1e0-a413-15cb558099b5.md)** object that is used as the default style for slicers. Read/write.|
| [DefaultTableStyle](2dc86b2c-0047-53b5-3cc4-af15c36b78cf.md)|Specifies the table style from the  **TableStyles** collection that is used as the default TableStyle. Read/write **Variant**.|
| [DefaultTimelineStyle](78261166-759a-8b18-c194-1f9124ca7e4e.md)|The name of the default slicer style of the workbook.  **Variant**. Read/Write|
| [DisplayDrawingObjects](78eec8af-416d-88e6-d1f4-0b97a008f752.md)|Returns or sets how shapes are displayed. Read/write  **Long**.|
| [DisplayInkComments](bce6b184-7640-f51c-1feb-1973de6ff739.md)|A  **Boolean** value that determines whether ink comments are displayed in the workbook. Read/write **Boolean**.|
| [DocumentInspectors](26d2575f-6e61-4509-6a67-45ae576bc9fe.md)|Returns a  ** [DocumentInspectors](8366d7cd-e016-bb99-d27f-749ca10352f1.md)** collection that represents the Document Inspector modules for the specified workbook. Read-only.|
| [DocumentLibraryVersions](b6338994-b5d9-ef9b-83b5-bdd47d0fd407.md)|Returns a  ** [DocumentLibraryVersions](075c0315-fade-6d45-9ab9-6c798f6f09ac.md)** collection that represents the collection of versions of a shared workbook that has versioning enabled and that is stored in a document library on a server.|
| [DoNotPromptForConvert](d2af6528-4d9f-6e94-4fa6-2322098b4b17.md)|Returns or sets if the user should be prompted to convert the workbook if the workbook contains features that are not supported by versions of Excel earlier than Excel 2007. Read/write  **Boolean**.|
| [EnableAutoRecover](04a82e4d-0231-adf1-1289-35514372c995.md)|Saves changed files, of all formats, on a timed interval. Read/write  **Boolean**.|
| [EncryptionProvider](13047af7-2e6e-6b64-05f1-8b4bd7a734ad.md)|Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Office Excel 2007 uses when encrypting documents. Read/write.|
| [EnvelopeVisible](d511a75a-ddd1-64f5-a09b-720657f64c09.md)| **True** if the e-mail composition header and the envelope toolbar are both visible. Read/write **Boolean**.|
| [Excel4IntlMacroSheets](70a8c8d0-1169-7c3d-904e-5a32a4693f45.md)|Returns a  ** [Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)**collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.|
| [Excel4MacroSheets](29161ab8-da75-c7b5-561d-f4423b8ab1ef.md)|Returns a  ** [Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)**collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.|
| [Excel8CompatibilityMode](8471493b-8733-cddf-75fa-42d3d1719300.md)|The  **Excel8CompatibilityMode** property provides developers with a way to check if the workbook is in compatibility mode. Read-only **Boolean**.|
| [FileFormat](ef722c3c-90ea-9810-b853-a3fff19d5c60.md)|Returns the file format and/or type of the workbook. Read-only  ** [XlFileFormat](4c0ebc4c-915c-c199-ee39-f4d14ba7b64e.md)**.|
| [Final](55d3a155-ca0c-1f7c-8612-80aac91a8eb3.md)|Returns or sets a  **Boolean** that indicates whether a workbook is final. Read/write **Boolean**.|
| [ForceFullCalculation](76f46d18-79e3-9828-d126-e221ae1a8157.md)|Returns or sets the specified workbook to forced calculation mode. Read/write.|
| [FullName](83f45d15-b009-f304-ca53-4daa80c06562.md)|Returns the name of the object, including its path on disk, as a string. Read-only  **String**.|
| [FullNameURLEncoded](589d98f7-e6fa-bc28-2c8f-7cb72009737a.md)|Returns a  **String** indicating the name of the object, including its path on disk, as a string. Read-only.|
| [HasPassword](e3cfdc90-1e82-5556-0064-e8269ba92539.md)| **True** if the workbook has a protection password. Read-only **Boolean**.|
| [HasVBProject](b4293266-40d9-a8a4-80ff-8b19ec7ed823.md)|Returns a  **Boolean** that represents whether a workbook has an attached Microsoft Visual Basic for Applications project. Read-only **Boolean**.|
| [HighlightChangesOnScreen](146f9a16-d32b-cc8f-fece-03864f0e13a2.md)| **True** if changes to the shared workbook are highlighted on-screen. Read/write **Boolean**.|
| [IconSets](c837d2a8-d21d-7432-a409-f49426368556.md)|This property is used to filter data in a workbook based on a cell icon from the  **IconSet** collection. Read-only.|
| [InactiveListBorderVisible](a6259862-9a29-f3a5-498f-633f51ec10e6.md)|A  **Boolean** value that specifies whether list borders are visible when a list is not active. Returns **True** if the border is visible. Read/write **Boolean**.|
| [IsAddin](b8c8b9f4-4be5-0260-957e-c6450f31a0c0.md)| **True** if the workbook is running as an add-in. Read/write **Boolean**.|
| [IsInplace](f492c09f-79d1-cde0-6cf1-db9644e41589.md)| **True** if the specified workbook is being edited in place. **False** if the workbook has been opened in Microsoft Excel for editing. Read-only **Boolean**.|
| [KeepChangeHistory](3dbc322e-2b93-ae3d-cb9e-64c657fc0f80.md)| **True** if change tracking is enabled for the shared workbook. Read/write **Boolean**.|
| [ListChangesOnNewSheet](77adf429-baa5-f2be-6139-c2b07dda5174.md)| **True** if changes to the shared workbook are shown on a separate worksheet. Read/write **Boolean**.|
| [Mailer](b020d3f6-7120-d03c-bc42-c297bcfbebf6.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| [Model](43ccdaa8-4a12-e745-88db-9db8a328ee5e.md)|Returns the top level  **Model** object which is the one Data Model for the workbook. Read-only|
| [MultiUserEditing](dc721463-ec34-8c52-6701-51c406beed23.md)| **True** if the workbook is open as a shared list. Read-only **Boolean**.|
| [Name](55526a99-da9c-7f14-55f7-dfe9bd8ff489.md)|Returns a  **String** value that represents the name of the object.|
| [Names](26be56ec-ea12-1600-602a-eb338d4a5a8b.md)|Returns a  ** [Names](ffecf89d-7bae-c470-8e37-608857a9de2a.md)** collection that represents all the names in the specified workbook (including all worksheet-specific names). Read-only **Names** object.|
| [Parent](4c039b5b-050f-8f4d-bc90-7982e92fb17c.md)|Returns the parent object for the specified object. Read-only.|
| [Password](5eaaf8cd-4344-946e-ecfa-c0f48946d2f2.md)|Returns or sets the password that must be supplied to open the specified workbook. Read/write  **String**.|
| [PasswordEncryptionAlgorithm](2745a8da-2a61-b949-115a-7f1112a0289e.md)|Returns a  **String** indicating the algorithm Microsoft Excel uses to encrypt passwords for the specified workbook. Read-only.|
| [PasswordEncryptionFileProperties](536ad729-424e-5f81-85e9-8a6ed71fb11a.md)| **True** if Microsoft Excel encrypts file properties for the specified password-protected workbook. Read-only **Boolean**.|
| [PasswordEncryptionKeyLength](2662f2f5-1ad0-4a75-82c0-3268f147948a.md)|Returns a  **Long** indicating the key length of the algorithm Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.|
| [PasswordEncryptionProvider](d5bcbbf2-8de9-6725-9cac-679d6c023b34.md)|Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.|
| [Path](f4cbf76a-2ed3-63b7-3262-45403d6f086e.md)|Returns a  **String** that represents the complete path to the workbook/file that this workbook object respresents.|
| [Permission](ef04f56e-a04d-c3d9-fdda-611be7bf9d39.md)|Returns a  **Permission** object that represents the permission settings in the specified workbook.|
| [PersonalViewListSettings](998320bf-d703-e42f-8b43-5a7b909a846d.md)| **True** if filter and sort settings for lists are included in the user's personal view of the shared workbook. Read/write **Boolean**.|
| [PersonalViewPrintSettings](6e4a0a9c-4eb0-d8e1-e9ce-8e9e618996b4.md)| **True** if print settings are included in the user's personal view of the shared workbook. Read-write **Boolean**.|
| [PivotTables](b11795e0-22c8-f089-c59a-5e3d7a09d5de.md)|Returns an object that represents a collection of all the PivotTable reports on a worksheet. Read-only.|
| [PrecisionAsDisplayed](4f0c8201-5b8d-5cb5-337c-944d2c7dd8d1.md)| **True** if calculations in this workbook will be done using only the precision of the numbers as they're displayed. Read/write **Boolean**.|
| [ProtectStructure](bf721b60-0ad1-f71c-7ef4-74d2196d320e.md)| **True** if the order of the sheets in the workbook is protected. Read-only **Boolean**.|
| [ProtectWindows](0f285fbe-2545-5c7d-9e3d-f08d57e78092.md)| **True** if the windows of the workbook are protected. Read-only **Boolean**.|
| [PublishObjects](b6418f80-5154-6e3f-7313-222e6438c0e1.md)|Returns the  ** [PublishObjects](33ad393e-5ab6-2531-5e5b-42930fc596c0.md)**collection. Read-only.|
| [ReadOnly](f3c0ec74-63af-ed76-f854-ce2382b9fcf3.md)| Returns **True** if the object has been opened as read-only. Read-only **Boolean**.|
| [ReadOnlyRecommended](3cae84e4-d5f0-f01c-64d9-ec586ffdf79c.md)| **True** if the workbook was saved as read-only recommended. Read-only **Boolean**.|
| [RemovePersonalInformation](f5cdc655-8ba9-6dd1-ab05-028d98c11972.md)| **True** if personal information can be removed from the specified workbook. The default value is **False**. Read/write  **Boolean**.|
| [Research](3a7ba740-314b-664b-3be6-1e8cdeded234.md)|Returns a  **Research** object that represents the research service for a workbook. Read-only.|
| [RevisionNumber](7ea9fde5-eb89-a9b0-b637-980f1533d8ec.md)|Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero). Read-only  **Long**.|
| [Saved](37eb8e08-2bfa-8065-2520-a71e291ab50c.md)| **True** if no changes have been made to the specified workbook since it was last saved. Read/write **Boolean**.|
| [SaveLinkValues](ee69911f-5a4a-5c2b-c14a-cd562f3ba9f4.md)| **True** if Microsoft Excel saves external link values with the workbook. Read/write **Boolean**.|
| [ServerPolicy](188f6c47-35e3-bb69-cb8d-9d78b5b8fea5.md)|Returns a  **ServerPolicy** object that represents a policy specified for a workbook stored on a server running SharePoint Server 2007 or later. Read-only.|
| [ServerViewableItems](2c10a647-2b2c-0640-9990-109b89444cd2.md)|Allows a developer to interact with the list of published objects in the workbook that are shown on the server. Read-only.|
| [SharedWorkspace](864fdee9-7149-360f-099d-e1a9b57a31db.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
| [Sheets](45e4e19e-55ea-9615-231d-9435ba6d5a63.md)|Returns a  ** [Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)** collection that represents all the sheets in the specified workbook. Read-only **Sheets** object.|
| [ShowConflictHistory](d8588b9e-3e4b-6224-aaa7-ce0b63ff0607.md)| **True** if the Conflict History worksheet is visible in the workbook that's open as a shared list. Read/write **Boolean**.|
| [ShowPivotChartActiveFields](8892b134-4882-e1ff-a265-65b36af66f1a.md)|This property controls the visibility of the PivotChart Filter Pane. Read/write  **Boolean**.|
| [ShowPivotTableFieldList](33c74c54-27ea-d230-c640-47109bdfd4a2.md)| **True** (default) if the PivotTable field list can be shown. Read/write **Boolean**.|
| [Signatures](b45f8036-c2d7-6113-e95c-ff78ee6a1f46.md)|Returns the digital signatures for a workbook. Read-only.|
| [SlicerCaches](1ebb7fd1-1742-815a-b4bb-4d25d6c9e705.md)|Returns the  ** [SlicerCaches](d6097f70-cdc7-3be7-575c-cf43a0765e10.md)** object associated with the workbook. Read-only.|
| [SmartDocument](19916b63-e93a-7f1e-532c-f4bbdb60622d.md)|Returns a  **SmartDocument** object that represents the settings for a smart document solution. Read-only.|
| [Styles](c9a70be9-cab5-ea5f-2e3f-949b1acf43d9.md)|Returns a  ** [Styles](146effdc-e007-814d-b110-f7bd944fc15f.md)**collection that represents all the styles in the specified workbook. Read-only.|
| [Sync](000c9739-13ab-d6eb-c1c3-2ce721911137.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
| [TableStyles](ac23db99-b2ce-0228-7808-728817736694.md)|Returns a  **TableStyles** collection object for the current workbook that refers to the styles used in the current workbook. Read-only.|
| [TemplateRemoveExtData](9851df1d-4e83-525a-8a43-bd84b0a94c74.md)| **True** if external data references are removed when the workbook is saved as a template. Read/write **Boolean**.|
| [Theme](1208f610-8c6f-9a62-3378-9566a7ee6b37.md)|Returns the theme applied to the current workbook. Read-only.|
| [UpdateLinks](c8d374d7-0b32-eb32-fa29-ab496d6786e7.md)|Returns or sets an  ** [XlUpdateLink](8ddd9876-7c24-09dd-5b89-33804adc2097.md)**constant indicating a workbook's setting for updating embedded OLE links. Read/write.|
| [UpdateRemoteReferences](055c1a88-c189-ddd3-c9b2-9458817cec90.md)| **True** if Microsoft Excel updates remote references in the workbook. Read/write **Boolean**.|
| [UserStatus](0df24f8a-b60b-fd8c-3436-903652487a09.md)|Returns a 1-based, two-dimensional array that provides information about each user who has the workbook open as a shared list. Read-only  **Variant**.|
| [UseWholeCellCriteria](b65093aa-37ca-2aa1-4f18-c90bc7536f74.md)| **True** if the workbook uses search patterns that match the entire content of a cell. Read-only **Boolean**. |
| [UseWildcards](92e7463c-6dbe-c409-461a-ca730402ad62.md)| **True** if the workbook enables wildcards for character string comparisons and searching. Read-only **Boolean**|
| [VBASigned](6e93161c-2fa4-1064-9b5d-a8eb96ad2bea.md)| **True** if the Visual Basic for Applications project for the specified workbook has been digitally signed. Read-only **Boolean**.|
| [VBProject](1bef5b7e-e169-fa4b-9810-6cd87ecd0a8d.md)|Returns a  **VBProject** object that represents the Visual Basic project in the specified workbook. Read-only.|
| [WebOptions](801742a2-f5d8-5311-ea24-fd428532ba80.md)|Returns the  ** [WebOptions](d573637f-1891-4602-c961-091795e47356.md)**collection, which contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page. Read-only.|
| [Windows](2352d6c9-720e-b58d-6e7c-049bf21a090d.md)|Returns a  ** [Windows](d5d0e3c9-9132-469c-d033-d29397dacd77.md)** collection that represents all the windows in the specified workbook. Read-only **Windows** object.|
| [Worksheets](8b7d660d-ca49-0bd0-dc57-64defa47bd5e.md)|Returns a  ** [Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)**collection that represents all the worksheets in the specified workbook. Read-only  **Sheets** object.|
| [WritePassword](ac89063e-6ef5-f7c5-abb0-4e6ef1c5fd05.md)|Returns or sets a  **String** for the write password of a workbook. Read/write.|
| [WriteReserved](96cc86d1-0e77-b6f3-3045-f6346de0f969.md)| **True** if the workbook is write-reserved. Read-only **Boolean**.|
| [WriteReservedBy](f053c197-3af3-9ab7-bee1-f72ee311a5b8.md)|Returns the name of the user who currently has write permission for the workbook. Read-only  **String**.|
| [XmlMaps](c7893167-bfa1-e1df-58f3-782b80322fad.md)| Returns an **XmlMaps** collection that represents the schema maps that have been added to the specified workbook. Read-only.|
| [XmlNamespaces](b93aba02-f831-6321-1c0d-2a30d417e57f.md)|Returns an  ** [XmlNamespaces](430f6773-2be5-8312-cd67-afb703ab0782.md)** collection that represents the XML namespaces contained in the specified workbook. Read-only.|
