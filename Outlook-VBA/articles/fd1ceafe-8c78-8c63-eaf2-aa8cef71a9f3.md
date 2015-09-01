
# SharingItem.Permission Property (Outlook)

 **Last modified:** July 28, 2015

Sets or returns an  ** [OlPermission](11126d37-33da-53f7-f5b6-ea8603998651.md)** constant that determines what permissions to grant the recipients on the ** [SharingItem](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)**. Read/write.

## Syntax

 _expression_. **Permission**

 _expression_A variable that represents a  **SharingItem** object.


## Remarks

The  **Permission** property should be synchronized with the ** [PermissionTemplateGuid](166c2975-b6be-d1ca-4aa8-ad7deb42c68d.md)** property to accurately reflect the permission status of the **SharingItem**. Setting the  **PermissionTemplateGuid** property to a valid GUID should also incur setting the **Permission** property to **OlPermission.olPermissionTemplate**.

 When no Information Rights Management (IRM) has been set up (in which case the **Permission** property is **OlPermission.olUnrestricted**), or the restriction is not to forward the  **SharingItem** (in which case the **Permission** property is **OlPermission.olDoNotForward**), the value of the  **PermissionTemplateGuid** property should be an empty string.

Although you can view content that is protected by IRM on any computer running the 2007 Microsoft Office system or a later version, you must have Microsoft Office Professional Edition 2003, Microsoft Office Outlook 2007, or a later version of Outlook to create or send an e-mail that is protected by IRM.


## See also


#### Concepts


 [SharingItem Object](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)
#### Other resources


 [SharingItem Object Members](719ad60e-2242-2c54-778f-006b61690389.md)
