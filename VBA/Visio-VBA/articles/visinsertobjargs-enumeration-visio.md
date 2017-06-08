---
title: VisInsertObjArgs Enumeration (Visio)
keywords: vis_sdr.chm70050
f1_keywords:
- vis_sdr.chm70050
ms.prod: visio
ms.assetid: 5057dcd2-388b-9b57-cbfe-e6f68376a392
ms.date: 06/08/2017
---


# VisInsertObjArgs Enumeration (Visio)

Flags to be passed to the  **Page.InsertObject** and **Page.InsertFromFile** methods.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visInsertAsControl**|8192|None.|
| **visInsertAsEmbed**|16384|None.|
| **visInsertDontShow**|4096|Don't execute the new object's show verb.|
| **visInsertIcon**|16|Display the new object as an icon|
| **visInsertLink**|8|If set, the new shape represents an OLE link to the named file. Otherwise, the InsertFromFile method produces an OLE object from the contents of the named file and embeds it in the document that contains the page, master, or group.|
| **visInsertNoDesignModeTransition**|256|If set, when an ActiveX control is inserted, prevents Microsoft Visio from transitioning to design mode. |

