
# RemoteItem Events (Outlook)
This object has the following events:

 **Last modified:** July 28, 2015


## Events



|**Name**|**Description**|
|:-----|:-----|
| [AfterWrite](806e9b23-9f08-6888-607a-4377af2c4d04.md)|Occurs after Microsoft Outlook has saved the item.|
| [AttachmentAdd](7cce4d2a-4071-9277-2cbb-5ebeba781f0a.md)|Occurs when an attachment has been added to an instance of the parent object.|
| [AttachmentRead](1a3a7f96-6d48-e93c-476b-2b06ee3807ef.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
| [AttachmentRemove](b31b2967-5014-1ced-67b7-4cc4793864e0.md)|Occurs when an attachment has been removed from an instance of the parent object.|
| [BeforeAttachmentAdd](03bee9f2-95cc-747a-c0fe-4d237b347cd9.md)|Occurs before an attachment is added to an instance of the parent object.|
| [BeforeAttachmentPreview](fcf508c5-280c-6b3c-d3db-eed7e8382cc2.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
| [BeforeAttachmentRead](739b8606-3e3a-1445-6355-896a6e897a6f.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  ** [Attachment](3e11582b-ac90-0948-bc37-506570bb287b.md)** object.|
| [BeforeAttachmentSave](bbccaae4-6e32-0e1a-0666-870dbfa1b678.md)|Occurs just before an attachment is saved.|
| [BeforeAttachmentWriteToTempFile](fb309e7f-b8a6-b73c-de7a-77a15a70249d.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
| [BeforeAutoSave](f33e1442-0e65-cc78-34ac-496b65ba565e.md)|Occurs before the item is automatically saved by Outlook.|
| [BeforeCheckNames](b34071cd-b43f-4801-b5da-6008eaef6ebf.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
| [BeforeDelete](0f1f4b6d-7a5a-2302-2b71-eea7bf7f1af9.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
| [BeforeRead](aa42bad1-3bab-a2f2-6565-9804dc90ae6d.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
| [Close](77276903-9e9e-713a-5844-c4efd36a020d.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
| [CustomAction](4c662153-6de7-8e6b-021c-f7f407e0d790.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
| [CustomPropertyChange](73d2e97b-eccd-d7ed-03e4-eb5e5fc345e3.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
| [Forward](f4af05e8-c0ea-915e-f49c-2470620e0735.md)|Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).|
| [Open](57094921-508c-7546-1981-3686bea7d325.md)|Occurs when an instance of the parent object is being opened in an  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)**. |
| [PropertyChange](630d4423-cb56-eef0-e1b1-1afe227c140d.md)|Occurs when an explicit built-in property (for example,  ** [Subject](57f0f242-6d04-175f-4ea2-25145787f5bd.md)**) of an instance of the parent object is changed. |
| [Read](78ad2650-7108-f617-6a04-74d7db8db4d7.md)|Occurs when an instance of the parent object is opened for editing by the user. |
| [ReadComplete](208867c1-b6dc-4ce8-e25a-13a8f6c686ca.md)|Occurs when Outlook has completed reading the properties of the item.|
| [Reply](47b49c1a-2e70-0265-d36d-58cf3800ffaf.md)|Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).|
| [ReplyAll](6616031a-7f71-bf18-5396-97707b1cccb1.md)|Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).|
| [Send](6b2ddae1-8732-c6d2-8dff-585118c3d051.md)|Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).|
| [Unload](8d105e1a-4923-4296-10b1-6e26fed51539.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
| [Write](a38eef6b-23da-ba10-ad94-cc63e2bf60c2.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  ** [Save](0f4e57ab-388c-7ba1-c6b8-f14bfc0ac73c.md)** or ** [SaveAs](1c2c7b68-5239-05f8-4291-d2584fe95194.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|
