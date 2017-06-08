---
title: UserPermission Members (Office)
ms.prod: office
ms.assetid: b9fdae9a-719b-9e1d-42aa-7553de91f9d1
ms.date: 06/08/2017
---


# UserPermission Members (Office)
Associates a set of permissions on the active document with a single user and an optional expiration date. Represents a member of the active document's  **Permission** collection.

Associates a set of permissions on the active document with a single user and an optional expiration date. Represents a member of the active document's  **Permission** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Remove](userpermission-remove-method-office.md)|Removes the specified  **UserPermission** object from the **[Permission](permission-object-office.md)** collection of the active document.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](userpermission-application-property-office.md)|Gets an  **Application** object that represents the container application for the **UserPermission** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Creator](userpermission-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **UserPermission** object was created. Read-only.|
|[ExpirationDate](userpermission-expirationdate-property-office.md)|Gets or sets the optional expiration date of the permissions on the active document assigned to the user associated with the specified  **UserPermission** object. Read/write.|
|[Parent](userpermission-parent-property-office.md)|Gets the  **Parent** object for the **UserPermission** object. Read-only.|
|[Permission](userpermission-permission-property-office.md)| Returns or sets a **MsoPermission** constant as a **Long** value representing the permissions on the active document assigned to the user associated with the specified **UserPermission** object. Read/write.|
|[UserId](userpermission-userid-property-office.md)|Gets the e-mail name of the user whose permissions on the active document are determined by the specified  **[UserPermission](userpermission-object-office.md)** object. Read-only.|

