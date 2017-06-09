---
title: Permission Members (Office)
ms.prod: office
ms.assetid: 75614d24-cd47-ef9b-aba5-112206daa358
ms.date: 06/08/2017
---


# Permission Members (Office)
The  **Permission** property of the **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **Permission** object.

The  **Permission** property of the **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **Permission** object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](permission-add-method-office.md)|Creates a set of permissions on the active document for the specified user. Returns a  **UserPermission** object.|
|[ApplyPolicy](permission-applypolicy-method-office.md)|Applies the specified permission policy to the active document.|
|[RemoveAll](permission-removeall-method-office.md)|Removes all  **UserPermission** objects from the **Permission** collection of the active document.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](permission-application-property-office.md)|Gets an  **Application** object that represents the container application for the **Permission** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Count](permission-count-property-office.md)|Gets a  **Long** indicating the number of items in the **Permission** object. Read-only.|
|[Creator](permission-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **Permission** object was created. Read-only.|
|[DocumentAuthor](permission-documentauthor-property-office.md)|Gets or sets the name in e-mail form of the author of the active document. Read/write.|
|[Enabled](permission-enabled-property-office.md)|Gets or sets a  **Boolean** value that indicates whether permissions are enabled on the active document. Read/write.|
|[EnableTrustedBrowser](permission-enabletrustedbrowser-property-office.md)|Gets or sets a value indicating whether to enable a browser from a trusted source. Read/write.|
|[Item](permission-item-property-office.md)|Gets a  **UserPermission** object that is a member of the **Permission** collection. The **UserPermission** object associates a set of permissions on the active document with a single user and an optional expiration date. Read-only.|
|[Parent](permission-parent-property-office.md)|Gets the  **Parent** object for the **Permission** object. Read-only.|
|[PermissionFromPolicy](permission-permissionfrompolicy-property-office.md)|Gets a  **Boolean** value that indicates whether a permission policy has been applied to the active document. Read-only.|
|[PolicyDescription](permission-policydescription-property-office.md)|Gets the description of the permissions policy applied to the active document. Read-only.|
|[PolicyName](permission-policyname-property-office.md)|Gets the name of the permissions policy applied to the active document. Read-only.|
|[RequestPermissionURL](permission-requestpermissionurl-property-office.md)|Gets or sets the file or Web site URL to visit or the e-mail address to contact for users who need additional permissions on the active document. Read/write.|
|[StoreLicenses](permission-storelicenses-property-office.md)|Gets or sets a  **Boolean** value that indicates whether the user's license to view the active document should be cached to allow offline viewing when the user cannot connect to a rights management server. Read/write.|

