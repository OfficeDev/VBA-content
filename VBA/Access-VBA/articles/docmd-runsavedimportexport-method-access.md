---
title: DoCmd.RunSavedImportExport Method (Access)
keywords: vbaac10.chm5878
f1_keywords:
- vbaac10.chm5878
ms.prod: access
api_name:
- Access.DoCmd.RunSavedImportExport
ms.assetid: cb0ade9a-5cd4-1225-5231-8266fdfb3690
ms.date: 06/08/2017
---


# DoCmd.RunSavedImportExport Method (Access)

Run a saved import or export specification.


## Syntax

 _expression_. **RunSavedImportExport**( ** _SavedImportExportName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SavedImportExportName_|Required|**Variant**| The name of a saved import or export specification to run.|

## Remarks

This method has the same effect as performing the following procedure in Access:


1. On the  **External Data** tab, click either **Saved Imports** or **Saved Exports**.
    
2. In the  **Manage Data Tasks** dialog box, on the **Saved Imports** or **Saved Exports** tab (depending on your choice in the preceding step), click the specification that you want to run.
    
3. Click  **Run**. 
    
Before running the  **RunSavedImportExport** method, make sure that the source and destination files exist, the source data is ready for importing, and that the operation will not accidentally overwrite any data in your destination file.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

