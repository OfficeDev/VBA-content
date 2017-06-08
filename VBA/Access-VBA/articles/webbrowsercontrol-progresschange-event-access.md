---
title: WebBrowserControl.ProgressChange Event (Access)
keywords: vbaac10.chm143142
f1_keywords:
- vbaac10.chm143142
ms.prod: access
api_name:
- Access.WebBrowserControl.ProgressChange
ms.assetid: 1a021887-6f0c-236a-2228-90a339407689
ms.date: 06/08/2017
---


# WebBrowserControl.ProgressChange Event (Access)

Occurs when the progress of a download operation is updated.


## Syntax

 _expression_. **ProgressChange**( ** _Progress_**, ** _ProgressMax_** )

 _expression_ A variable that represents a **WebBrowserControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Progress_|Required|**Long**|Specifies the amount of total progress to show, or -1 when progress is complete.|
| _ProgressMax_|Required|**Long**|Specifies the maximum progress value. |

### Return Value

nothing


## Remarks

You can use the information provided by this event to display the number of bytes downloaded or to update a progress indicator.

To calculate the percent of progress to show in a progress indicator, multiply the value of Progress by 100, and divide by the value of  _ProgressMax_; unless _Progress_ is -1, in which case the container indicates that the operation is finished or hides the progress indicator.


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

