---
title: MailMerge.ShowWizardEx Method (Publisher)
keywords: vbapb10.chm6225944
f1_keywords:
- vbapb10.chm6225944
ms.prod: publisher
api_name:
- Publisher.MailMerge.ShowWizardEx
ms.assetid: 3815204f-5f09-5a25-a2e4-5de4889c9919
ms.date: 06/08/2017
---


# MailMerge.ShowWizardEx Method (Publisher)

Displays the specified catalog or mail merge wizard in a document.


## Syntax

 _expression_. **ShowWizardEx**( **_ShowDocumentStep_**,  **_ShowTemplateStep_**,  **_ShowDataStep_**,  **_ShowWriteStep_**,  **_ShowPreviewStep_**,  **_ShowMergeStep_**, **_MergeType_**, **_iStep_**)

 _expression_A variable that represents a  **MailMerge** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ShowDocumentStep|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions,  **True** (the default) displayed the "Select a merge type" step. **False** removed the step.|
|ShowTemplateStep|Optional| **Boolean**| This parameter does not apply to Microsoft Publisher.|
|ShowDataStep|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions,  **True** (the default) displayed the "Select data source" step. **False** removed the step.|
|ShowWriteStep|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions,  **True** (the default) displayed the "Create your publication" step. **False** removed the step.|
|ShowPreviewStep|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions,  **True** (the default) displayed the "Preview your publication" step. **False** removed the step.|
|ShowMergeStep|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions,  **True** (the default) displayed the "Complete the merge" step. **False** removed the step.|
|MergeType|Optional| **PbMergeType**|The merge type to use. See Remarks for possible values.|
|iStep|Optional| **Long**|The initial step. See Remarks for information about default values.|

## Remarks

The MergeType parameter can be one of the  **[PbMergeType](pbmergetype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library. The default is **pbMergeDefault**.

Passing  **pbMergeDefault** for MergeType starts a new mail merge; if the publication is already a merge, it leaves the merge type unchanged.

Passing a merge type that is different from the current publication's merge type changes the publication to that new type of merge, but disconnects the data source. Doing so results in the loss of previously inserted fields when the change is to or from a catalog merge type.

Wizard steps correspond to the sequence of merge task panes in the user interface. If no data source is connected, the merge wizard always starts on the first step (the first task pane). If a data source is connected, the wizard starts on Step 2 by default, unless you use the iStep parameter to specify starting with Step 1 or Step 3.


## Example

This example checks whether the  **Mail Merge Wizard** is closed, and if it is, displays it.


```vb
Public Sub ShowWizardEx_Example() 
 With ActiveDocument.MailMerge 
 If .WizardState = 0 Then 
 .ShowWizardEx 
 End If 
 End With 
End Sub
```


