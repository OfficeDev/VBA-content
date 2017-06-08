---
title: Presentation.RemoveDocumentInformation Method (PowerPoint)
keywords: vbapp10.chm583094
f1_keywords:
- vbapp10.chm583094
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.RemoveDocumentInformation
ms.assetid: 2c9d5cc5-8fc9-d650-b1cf-9fa3e409be1c
ms.date: 06/08/2017
---


# Presentation.RemoveDocumentInformation Method (PowerPoint)

Removes document information, such as personal information, comments, and document properties, from a Microsoft PowerPoint presentation.


## Syntax

 _expression_. **RemoveDocumentInformation**( **_Type_** )

 _expression_ An expression that returns a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**PpRemoveDocInfoType**|Type of information to be removed.|

## Remarks

The  _Type_ parameter value can be a combination of one or more of these **PpRemoveDocInfoType** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**ppRDIAll**|Remove all document information.|
|**ppRDIComments**|Remove comments.|
|**ppRDIContentType**|Remove content type information.|
|**ppRDIDocumentManagementPolicy**|Remove document management policy information.|
|**ppRDIDocumentProperties**|Remove document properties.|
|**ppRDIDocumentServerProperties**|Remove document server properties.|
|**ppRDIDocumentWorkspace**|Remove document workspace information.|
|**ppRDIInkAnnotations**|Remove Ink annotations.|
|**ppRDIPublishPath**|Remove publication path information.|
|**ppRDIRemovePersonalInformation**|Remove personal information.|
|**ppRDISlideUpdateInformation**|Remove slide update information.|

## Example

The following example shows how to use the  **RemoveDocumentInformation** method to remove comments and Ink annotations from the active presentation.


```vb
Public Sub RemoveDocumentInformation_Example()



    ActivePresentation.RemoveDocumentInformation (ppRDIComments + ppRDIInkAnnotations)



End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

