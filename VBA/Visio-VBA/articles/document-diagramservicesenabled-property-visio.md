---
title: Document.DiagramServicesEnabled Property (Visio)
keywords: vis_sdr.chm10562435
f1_keywords:
- vis_sdr.chm10562435
ms.prod: visio
api_name:
- Visio.DiagramServicesEnabled
ms.assetid: 1a492029-31c8-85bb-0843-31c0a1200055
ms.date: 06/08/2017
---


# Document.DiagramServicesEnabled Property (Visio)

Determines which, if any, diagram services are enabled for the document. Read/write.


## Syntax

 _expression_ . **DiagramServicesEnabled**

 _expression_ An expression that returns a **[Document](document-object-visio.md)** object.


### Return Value

 **Long**


## Remarks

Visio has several diagram behaviors, including structured-diagram behaviors and AutoSize behaviors. Structured-diagram behaviors define when container-membership relationships and callout associations are created. AutoSize behaviors define when Visio automatically resizes the drawing page to adjust to changes in its contents.

In your solution, you can take advantage of these new diagram behaviors by using the  **DiagramServicesEnabled** property to enable the services that aggregate these behaviors. When your solution modifies the diagram, Visio invokes the diagram behaviors associated with any of the services that are currently enabled.

The value of the  **DiagramServicesEnabled** property setting must be one or a bitwise combination of the following constants from the **[VisDiagramServices](visdiagramservices-enumeration-visio.md)** enumeration.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visServiceNone**|0|No diagram services.|
| **visServiceAll**|-1|All diagram services.|
| **visServiceAutoSizePage**|1|AutoSize (automatic page-sizing) behaviors.|
| **visServiceStructureBasic**|2|Structured-diagram behaviors that maintain existing relationships but do not create new relationships.|
| **visServiceStructureFull**|4|Structured-diagram behaviors that match all those in the user interface (UI).|
| **visServiceVersion140**|7|All diagram services that exist in Visio.|
| **visServiceVersion150**|8|All diagram services that exist in Visio.|
 If you combine **visServiceStructureBasic** and **visServiceStructureFull** , the latter overrides the former. However, you can combine **visServiceAutoSizePage** with either **visServiceStructureBasic** (3) or **visServiceStructureFull** (5) and assign either of those values to the property.

Diagram services apply only to solutions that manipulate Visio programmatically (by Automation). They do not have any effect on the behaviors that are exposed in the UI. UI settings that disable these behaviors have no effect on behaviors that are triggered programmatically.

By default, diagram services are disabled for a document. You must enable any services you want to take advantage of before your solution modifies the diagram. Diagram service settings are not persisted from one session of Visio to the next.


