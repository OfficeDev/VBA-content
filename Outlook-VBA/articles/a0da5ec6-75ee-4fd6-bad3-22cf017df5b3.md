
# DocumentItem Events (Outlook)
This object has the following events:

 **Last modified:** July 28, 2015


## Events



|**Name**|**Description**|
|:-----|:-----|
| [AfterWrite](f810f61f-9fad-6001-d9fa-389ce4003ac7.md)|Occurs after Microsoft Outlook has saved the item.|
| [AttachmentAdd](229bc1b9-64bb-2198-1ec9-10f7129a59b9.md)|Occurs when an attachment has been added to an instance of the parent object.|
| [AttachmentRead](46cb82e1-1705-acc1-6bc3-e673ed2be44a.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
| [AttachmentRemove](c921bdd1-f922-8cd4-a31c-fd880b447099.md)|Occurs when an attachment has been removed from an instance of the parent object.|
| [BeforeAttachmentAdd](cd440e8a-c79a-d1b4-9d03-940b2f3fa50b.md)|Occurs before an attachment is added to an instance of the parent object.|
| [BeforeAttachmentPreview](687c0c41-c423-a30f-3fb6-562c2ab76f0c.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
| [BeforeAttachmentRead](22ed23a8-42a5-09bd-73b9-10591bfa7de9.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  ** [Attachment](3e11582b-ac90-0948-bc37-506570bb287b.md)** object.|
| [BeforeAttachmentSave](554f3e7d-9757-c044-2cfd-56614be6b27b.md)|Occurs just before an attachment is saved.|
| [BeforeAttachmentWriteToTempFile](09ec6f62-e5c6-1884-ba77-e4865978d0ba.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
| [BeforeAutoSave](3aaf57a3-bcc2-d0ba-6fd9-d801452dc4ca.md)|Occurs before the item is automatically saved by Outlook.|
| [BeforeCheckNames](0798f1bc-4a7e-7f85-0719-31f5f937cfc3.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
| [BeforeDelete](73900e17-571c-e972-eeca-fb0d591a4641.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
| [BeforeRead](5b494a75-3d56-ee3f-8415-b44bca720440.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
| [Close](13aecc0c-9e71-7e47-147a-0af020c857bd.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
| [CustomAction](eec2389c-45bf-38fb-46fe-c319cac12319.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
| [CustomPropertyChange](11fc60a4-39ef-3e39-d9af-0a5ccf3cbc43.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
| [Forward](394f3c85-61b8-4f2e-a64a-d2f61f42c6f4.md)|Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).|
| [Open](e7d95148-9fa2-3f0f-cbfc-f835c9017c3b.md)|Occurs when an instance of the parent object is being opened in an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)**. |
| [PropertyChange](ec757f98-db44-585e-1a4a-5b3044428dec.md)|Occurs when an explicit built-in property (for example,  ** [Subject](57f0f242-6d04-175f-4ea2-25145787f5bd.md)**) of an instance of the parent object is changed. |
| [Read](da5e82e6-43b9-d040-e529-2388049a8e1b.md)|Occurs when an instance of the parent object is opened for editing by the user. |
| [ReadComplete](5a47b0f4-dfa9-9cf6-8efa-7ab45c1f90d7.md)|Occurs when Outlook has completed reading the properties of the item.|
| [Reply](2a35c8d0-5d84-35cf-3ee2-4bbbf053428e.md)|Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).|
| [ReplyAll](b60ee051-6fb7-3572-e359-57093495adb2.md)|Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).|
| [Send](7e77c1c3-f6dd-13d1-ed76-b37e7dd6e82a.md)|Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).|
| [Unload](e634c3f3-e637-f18c-0f7e-2e5cb18566a3.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
| [Write](ae8c445f-cf46-9544-7073-bf08638b9247.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  ** [Save](0cb1716d-6e53-6188-0feb-3c4ece9ab0a6.md)** or ** [SaveAs](b9264e62-1302-617f-4c9d-74844c96a38d.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|
