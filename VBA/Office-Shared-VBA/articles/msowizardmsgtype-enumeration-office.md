---
title: MsoWizardMsgType Enumeration (Office)
ms.prod: office
api_name:
- Office.MsoWizardMsgType
ms.assetid: af88d063-45c9-8bf6-2707-dc27df02d3bb
ms.date: 06/08/2017
---


# MsoWizardMsgType Enumeration (Office)

Specifies context under which a wizard's callback procedure is called. Used as an argument in a callback procedure designed for use with a custom wizard.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**msoWizardMsgLocalStateOff**|2|User clicked the right button in the decision or branch balloon.|
|**msoWizardMsgLocalStateOn**|1|Not supported.|
|**msoWizardMsgResuming**|5|Passed to the  **ActivateWizard** method if **msoWizardActResume** is specified for the Act argument.|
|**msoWizardMsgShowHelp**|3|User clicked the left button in the decision or branch balloon.|
|**msoWizardMsgSuspending**|4|Passed to the  **ActivateWizard** method if **msoWizardActSuspend** is specified for the Act argument.|

