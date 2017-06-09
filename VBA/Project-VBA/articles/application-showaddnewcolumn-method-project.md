---
title: Application.ShowAddNewColumn Method (Project)
keywords: vbapj.chm709
f1_keywords:
- vbapj.chm709
ms.prod: project-server
api_name:
- Project.Application.ShowAddNewColumn
ms.assetid: 2f13b46a-da46-453d-1165-f9a1d9b06377
ms.date: 06/08/2017
---


# Application.ShowAddNewColumn Method (Project)

Shows or hides the  **Add New Column** column at the right side of the active sheet view.


## Syntax

 _expression_. **ShowAddNewColumn**( ** _Show_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if the **Show 'Add New Column' interface** option is selected. **False** if the option is cleared. The default value is **True**.|

### Return Value

 **Boolean**


## Remarks

The  **ShowAddNewColumn** method does not apply to views that do not use tables, such as the following: Network Diagram (PERT chart), Task Entry, Resource Entry, Calendar, or Timeline.

If a view uses a table, you can set individual views to show the  **Add New Column** column. To open the **Table Definition** dialog box for a view, do the following on the **VIEW** ribbon:


1. In the  **Other Views** drop-down list, open the **More Views** dialog box, and then edit the view to find the table that the view uses. For example, the Task Usage view uses the Usage table.
    
2. Close the  **View Definition** dialog box and the **More Views** dialog boxes.
    
3. In the  **Tables** drop-down list, open the **More Tables** dialog box, select the table, and then click **Edit**.
    
4. Select or clear the  **Show 'Add New Column' interface** option in the **Table Definition** dialog box.
    

