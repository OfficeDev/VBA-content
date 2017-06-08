---
title: Workbook.CheckInWithVersion Method (Excel)
keywords: vbaxl10.chm199238
f1_keywords:
- vbaxl10.chm199238
ms.prod: excel
api_name:
- Excel.Workbook.CheckInWithVersion
ms.assetid: 3b37cea5-8795-bcbb-9c4b-d30b2b9a095e
ms.date: 06/08/2017
---


# Workbook.CheckInWithVersion Method (Excel)

Saves a workbook to a server from a local computer, and sets the local workbook to read-only so that it cannot be edited locally.


## Syntax

 _expression_ . **CheckInWithVersion**( **_SaveChanges_** , **_Comments_** , **_MakePublic_** , **_VersionType_** )

 _expression_ A variable that returns a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**| **True** to save the workbook to the server location. The default is **True** .|
| _Comments_|Optional| **Variant**|Comments for the revision of the workbook being checked in (applies only if  _SaveChanges_ is set to **True** ).|
| _MakePublic_|Optional| **Variant**| **True** to allow the user to publish the workbook after it is checked in.|
| _VersionType_|Optional| **Variant**|Specifies versioning information for the workbook. |

### Return Value

Nothing


## Remarks

Setting the  _MakePublic_ parameter to **True** submits the workbook for the approval process, which can eventually result in a version of the workbook being published to users with read-only rights to the workbook (applies only if _SaveChanges_ is set to **True** ).

To take advantage of the collaboration features built into Microsoft Excel, documents must be stored on a Microsoft SharePoint Server. 


## Example

The following example uses the  **[CanCheckIn](workbook-cancheckin-method-excel.md)** method to determine whether the workbook has been stored on a Microsoft SharePoint Server. If the workbook has been stored on a server, the example calls the **CheckInWithVersion** method to check in the workbook along with the specified comments and version number, save changes to the server location, and submit the workbook for the approval process.

This example is for a workbook-level customization.




```vb
Private Sub WorkbookCheckIn() 
 If ActiveWorkbook.CanCheckIn Then 
 ActiveWorkbook.CheckInWithVersion _ 
 True, _ 
 "My updates.", _ 
 True, _ 
 XlCheckInVersionType.xlCheckInMinorVersion 
 Else 
 MessageBox.Show ("This workbook cannot be checked in") 
 End If 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

