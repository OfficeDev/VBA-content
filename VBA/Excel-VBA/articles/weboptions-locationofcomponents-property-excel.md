---
title: WebOptions.LocationOfComponents Property (Excel)
keywords: vbaxl10.chm662081
f1_keywords:
- vbaxl10.chm662081
ms.prod: excel
api_name:
- Excel.WebOptions.LocationOfComponents
ms.assetid: 0581343b-e93e-1413-4348-529f48a166eb
ms.date: 06/08/2017
---


# WebOptions.LocationOfComponents Property (Excel)

Returns or sets the central URL (on the intranet or Web) or path (local or network) to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. The default value is the local or network installation path for Microsoft Office. Read/write  **String** .


## Syntax

 _expression_ . **LocationOfComponents**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

Office Web components are automatically downloaded with the specified Web page if the  **[DownloadComponents](weboptions-downloadcomponents-property-excel.md)** property is is set to **True** , the components are not already installed, the path is valid and points to a location that contains the necessary components, and the user has a valid Microsoft Office license.


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

