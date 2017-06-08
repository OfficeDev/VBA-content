---
title: Signature Members (Office)
ms.prod: office
ms.assetid: 1054db23-fe1c-f81f-e44b-d8c2c82ca7fa
ms.date: 06/08/2017
---


# Signature Members (Office)
Represents a digital signature attached to a document.  **Signature** objects are contained in the **SignatureSet** collection of the **Document** object.

Represents a digital signature attached to a document.  **Signature** objects are contained in the **SignatureSet** collection of the **Document** object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](signature-delete-method-office.md)|Deletes the  **Signature** object from the collection.|
|[ShowDetails](signature-showdetails-method-office.md)|Displays details related to a signature packet.|
|[Sign](signature-sign-method-office.md)|Creates a signature packet.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](signature-application-property-office.md)|Gets an  **Application** object that represents the container application for the **Signature** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[CanSetup](signature-cansetup-property-office.md)|Gets a  **Boolean** value indicating whether the user can set properties of the **Signature** object. Read-only.|
|[Creator](signature-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **Signature** object was created. Read-only.|
|[Details](signature-details-property-office.md)|Gets information about a signature. Read-only.|
|[IsSignatureLine](signature-issignatureline-property-office.md)|Gets a value indicating whether this is a signature line. Read-only.|
|[IsSigned](signature-issigned-property-office.md)|Gets a  **Boolean** value indicating whether the document was signed successfully. Read-only.|
|[Parent](signature-parent-property-office.md)|Gets the  **Parent** object for the Signature object. Read-only.|
|[Setup](signature-setup-property-office.md)|Gets a  **SignatureSetup** object that provides access to various properties of a signature packet. Read-only.|
|[SignatureLineShape](signature-signaturelineshape-property-office.md)|Gets the  **Shape** object associated with a **Signature** object that is a signature line. Read-only.|
|[SortHint](signature-sorthint-property-office.md)|Gets a value representing the sort order of the signatures in a packet with multiple signatures. Read-only.|

