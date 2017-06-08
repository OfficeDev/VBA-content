---
title: Document.Mode Property (Visio)
keywords: vis_sdr.chm10513925
f1_keywords:
- vis_sdr.chm10513925
ms.prod: visio
api_name:
- Visio.Document.Mode
ms.assetid: 40ebcc64-43dc-79f4-2802-9cd9dba633ab
ms.date: 06/08/2017
---


# Document.Mode Property (Visio)

Determines whether a document is in run mode or design mode. Read/write.


## Syntax

 _expression_ . **Mode**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisDocModeArgs


## Remarks

A Microsoft Visio document is either in run mode or in design mode, just as a Microsoft Visual Basic form is either running or being designed.

The following are the fundamental distinctions between run mode and design mode:


- ActiveX controls hosted in a document are told not to fire events when the document is in design mode, and to fire events when in run mode.
    
- Visio doesn't source events from any object whose document is in design mode.
    
The run/design mode of a Visio document is reported in the Visio user interface by the  **Design Mode** control on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab. The appearance of this control is the same as that of the **Design Mode** button in the Visual Basic Editor window. If the control appears pressed, the document (project) is in design mode. If it does not appear pressed, the document (project) is in run mode.

The run/design mode of a Visio document is synchronized with the run/design state of the document's Visual Basic for Applications (VBA) project, provided the document has a project. If the document transitions to or from run mode, the project's mode switches, and vice versa. This means that if code in a document's project sets the document's mode to design mode ( **ThisDocument.Mode** = **visDocModeDesign** ), the project in which the code runs transitions to design mode and any statements following the mode-assignment statement are not executed. However, code in a document can put another document (project) into design mode and keep running.

A document's mode is not a persistent property. By default, a Visio document opens in design mode unless the document is from a trusted publisher, is digitally signed, or is in a trusted location. A document that meets one of these criteria opens in run mode.

However, you can change default settings in the  **Macro Settings** category of the Visio **Trust Center** (Click the **File** tab, click **Options**, click  **Trust Center**, and then click  **Trust Center Settings**.) If  **Disable all macros except digitally signed macros** is selected, Visio documents not in a trusted location open in run mode only if they are digitally signed. If **Disable all macros without notification** or **Disable all macros with notification** is selected, documents not in a trusted location open in design mode. If **Enable all macros** is selected, documents always open in run mode, but this option presents a security risk and is not recommended.


