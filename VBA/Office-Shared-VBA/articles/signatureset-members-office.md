---
title: SignatureSet Members (Office)
ms.prod: office
ms.assetid: abe810a3-ffe4-ee26-8df7-d68cfbf3bf1e
ms.date: 06/08/2017
---


# SignatureSet Members (Office)
A collection of  **Signature** objects that correspond to the digital signature attached to a document.

A collection of  **Signature** objects that correspond to the digital signature attached to a document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddNonVisibleSignature](signatureset-addnonvisiblesignature-method-office.md)|Creates a signature packet when digitally signing a document.|
|[AddSignatureLine](signatureset-addsignatureline-method-office.md)|Adds lines to a document where signatures are collected.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](signatureset-application-property-office.md)|Gets an  **Application** object that represents the container application for the **SignatureSet** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[CanAddSignatureLine](signatureset-canaddsignatureline-property-office.md)|Gets a  **Boolean** value indicating whether you can add a signature line to a document. Read-only.|
|[Count](signatureset-count-property-office.md)|Gets a  **Long** indicating the number of items in the **SignatureSet** object. Read-only.|
|[Creator](signatureset-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **SignatureSet** object was created. Read-only.|
|[Item](signatureset-item-property-office.md)|Gets a  **Signature** object that corresponds to one of the digital signatures with which the document is currently signed. Read-only.|
|[Parent](signatureset-parent-property-office.md)|Gets the  **Parent** object for the **SignatureSet** object. Read-only.|
|[ShowSignaturesPane](signatureset-showsignaturespane-property-office.md)|Gets or sets a  **Boolean** value indicating whether the **Signature** task pane should be displayed. Read/write.|
|[Subset](signatureset-subset-property-office.md)|Gets or sets a value that acts as a filter on the available  **Signature** objects for a document. Read/write.|

