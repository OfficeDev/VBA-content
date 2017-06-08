---
title: InvisibleApp.PromptForSummary Property (Visio)
keywords: vis_sdr.chm17514170
f1_keywords:
- vis_sdr.chm17514170
ms.prod: visio
api_name:
- Visio.InvisibleApp.PromptForSummary
ms.assetid: 932d890d-0c45-fe74-57f8-03662b9c9def
ms.date: 06/08/2017
---


# InvisibleApp.PromptForSummary Property (Visio)

Determines whether Microsoft Visio prompts for document properties when it saves a document. Read/write.


## Syntax

 _expression_ . **PromptForSummary**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Integer


## Remarks

This property corresponds to the  **Prompt for document properties on first save** check box on the **Save** tab in the **Visio Options** dialog box (click the **File** tab and then click **Options**).


## Example

This Microsoft Visual Basic for Applications (VBA) macro switches the  **PromptForSummary** property of the Visio **Application** object.


```vb
 
Public Sub PromptForSummary_Example() 
  
    Application.PromptForSummary = Not Application.PromptForSummary  
 
End Sub
```


