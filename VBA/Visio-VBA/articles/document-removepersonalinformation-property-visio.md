---
title: Document.RemovePersonalInformation Property (Visio)
keywords: vis_sdr.chm10551675
f1_keywords:
- vis_sdr.chm10551675
ms.prod: visio
api_name:
- Visio.Document.RemovePersonalInformation
ms.assetid: b05eb59e-9906-10f9-8819-60f8f0f1d4f5
ms.date: 06/08/2017
---


# Document.RemovePersonalInformation Property (Visio)

Determines if personal information about a file is saved when the user saves the file in Microsoft Visio. Read/write.


## Syntax

 _expression_ . **RemovePersonalInformation**

 _expression_ A variable that represents a **Document** object.


### Return Value

Boolean


## Remarks

The personal information that can be saved with a file appears on the  **Summary** tab of the **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**). Alternatively, you can use various Automation properties, for example the  **Document.Creator** property, to set this information.

Setting the  **RemovePersonalInformation** property is equivalent to setting the **Remove personal information from file properties on save** option under **Document-specific settings** on the **Privacy Options** page of the **Trust Center** dialog box (click the **File** tab, click **Options**, click  **Trust Center**, and then click  **Trust Center Settings**).


