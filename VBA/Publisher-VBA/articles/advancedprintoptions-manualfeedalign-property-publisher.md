---
title: AdvancedPrintOptions.ManualFeedAlign Property (Publisher)
keywords: vbapb10.chm7077928
f1_keywords:
- vbapb10.chm7077928
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.ManualFeedAlign
ms.assetid: 5c2dc0a7-981f-731d-6a85-0971c7e19a62
ms.date: 06/08/2017
---


# AdvancedPrintOptions.ManualFeedAlign Property (Publisher)

Gets or sets the alignment (left, right, or center) of where envelopes are fed to the printer's manual feed. Read/write.


## Syntax

 _expression_. **ManualFeedAlign**

 _expression_A variable that represents an  **AdvancedPrintOptions** object.


### Return Value

 **PbPlacementType**


## Remarks

The  **ManualFeedAlign** property setting, in conjunction with the ** [AdvancedPrintOptions.ManualFeedDirection](advancedprintoptions-manualfeeddirection-property-publisher.md)** property setting, corresponds to the **Envelope feed method** setting in the **Envelope Setup** dialog box in the Microsoft Publisher user interface. (On the **File** menu, click **Print Setup**. On the  **Printer Details** tab, click **Advanced Printer Setup**. On the  **Printer Setup Wizard** tab, click **Envelope Setup Dialog**).

Possible values for  **ManualFeedAlign** are **pbPlacementCenter** (3), **pbPlacementLeft** (1), and **pbPlacementRight** (2).


## See also


#### Concepts


 [AdvancedPrintOptions Object](advancedprintoptions-object-publisher.md)

