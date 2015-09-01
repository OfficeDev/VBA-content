
# Categorize Your Outlook Items

 **Last modified:** July 28, 2015

Microsoft Outlook provides color categorization functionality, in which Outlook items can be categorized and displayed by category. Multiple color categories can be applied to a single Outlook item, and Outlook items can be grouped or sorted by color category. Shortcut keys can be assigned to each color category, to allow users to more easily categorize items. Color categories are user-defined, and can be created, deleted, and changed either programmatically or by user action within the Outlook user interface.

The  ** [Category](143ef095-54b0-cbe2-e356-632029061ac2.md)** object represents a single user-defined color category in the Master Category List, the list of color categories presented in the Outlook user interface and represented by the ** [Categories](3963afca-3a7e-38d7-1347-7e1467be3a10.md)** collection of the ** [NameSpace](f0dcaa19-07f5-5d42-a3bf-2e42b7885644.md)** object. **Category** objects are identified with a globally unique identifier (GUID) when created, and this identifier cannot be changed. However, the name, color, and shortcut key associated with a color category can be changed by setting the ** [Name](b9a711e9-f79d-f4f7-88bb-eaeb61d64089.md)**,  ** [Color](42814031-97ee-bb71-7c24-4ddd367d793c.md)**, and  ** [ShortcutKey](c78f882a-ab02-5218-e71f-362c86b4dfe1.md)** properties, respectively, of the **Category** object. The ** [CategoryID](e75ed17a-940f-2325-8739-1367329854d2.md)** property can be used to retrieve the identifier of a **Category** object.


## Assigning Categories to Outlook Items

Categories can be assigned to Outlook items by specifying the names of the appropriate  **Category** objects in a comma-delimited string in the **Categories** property of the following objects:



| ** [AppointmentItem](204a409d-654e-27aa-643a-8344c631b82d.md)**| ** [RemoteItem](6302aaff-cdcf-4d86-60f1-4bed15540d9f.md)**|
| ** [ContactItem](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)**| ** [ReportItem](16ebe336-72e0-42f6-99d3-edecc3ea284d.md)**|
| ** [DistListItem](027c3986-abff-d9b1-ecc2-26d60805e952.md)**| ** [SharingItem](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)**|
| ** [DocumentItem](7b0a6af0-6632-3ff6-841f-5b081d0d68d8.md)**| ** [PostItem](de44065d-4e93-315a-279f-7b92f09c0465.md)**|
| ** [JournalItem](6e850295-39f9-47b8-e866-9622e9958c69.md)**| ** [TaskItem](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)**|
| ** [MailItem](14197346-05d2-0250-fa4c-4a6b07daf25f.md)**| ** [TaskRequestAcceptItem](a2905f72-0a67-b07d-7f85-84fe4de17c25.md)**|
| ** [MeetingItem](b75730f5-b395-3d66-5acd-b64fd8fcd78f.md)**| ** [TaskRequestDeclineItem](e842c7c0-7943-9219-329b-30b892ab99b0.md)**|
| ** [MobileItem](http://msdn.microsoft.com/library/da8149d5-66d3-ea02-941f-e7f2f9eb6bc3%28Office.15%29.aspx)**| ** [TaskRequestItem](2908a28a-634c-e786-aa53-f3e32038b727.md)**|
| ** [NoteItem](ddf5baaa-6e13-a6fb-96e8-311e7761fa98.md)**| ** [TaskRequestUpdateItem](5bc407fe-b3f6-3e46-8b91-e2ed96292cec.md)**|
Outlook items are displayed based on the category name stored in the  **Categories** property of that Outlook item. Because category names are stored as part of the Outlook item, it is possible to have a category name in an Outlook item that is not present in the Master Category List. For example, a category may have been removed.

If a  **Category** object with a corresponding **Name** property value does not exist in the **Categories** collection of the **NameSpace** object that contains the Outlook item, the category name associated with that Outlook item is still displayed, but without an associated color.

