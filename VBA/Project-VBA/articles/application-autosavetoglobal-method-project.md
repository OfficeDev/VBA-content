---
title: Application.AutoSaveToGlobal Method (Project)
keywords: vbapj.chm1500
f1_keywords:
- vbapj.chm1500
ms.prod: project-server
api_name:
- Project.Application.AutoSaveToGlobal
ms.assetid: 8b8d0169-a1c1-8771-bc90-503a17e00b26
ms.date: 06/08/2017
---


# Application.AutoSaveToGlobal Method (Project)

Specifies whether to automatically add new views, field templates, filters, and groups to the global template (Global.mpt).


## Syntax

 _expression_. **AutoSaveToGlobal**( ** _OnOff_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OnOff_|Optional|**Variant**|If  **True**, automatically save. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

If  **AutoSaveToGlobal** is off, you can manually save views, groups, and other items to the global template by using the **Organizer** dialog box. Click the **Office Button**, click the  **Info** tab, and then click **Manage Global Template**.


 **Note**  If  **AutoSaveToGlobal** successfully runs, it always returns **True**.

To see the results, run  `AutoSaveToGlobal OnOff:=True` in the **Immediate** pane of the VBE, and then create and save a view. For example, do the following:


1. In a new project, create three tasks (T1, T2, and T3) and two resources (R1 and R2).
    
2. Assign one of the tasks to R1 and the other two tasks to R2.
    
3. Click the  **View** tab in the Ribbon. In the **Data** group, click **Using Resource** in the drop-down list for **Filter**. 
    
4. In the  **Using Resource** dialog box, select R2 for the task filter.
    
5. In the  **Resource Views** group, click **Other Views**, and then click  **Save View**. For example, save the view with the name  **R2 View Test**.
    
6. Click  **Other Views** again, and then click **More Views**. The  **Views** list contains the view you saved.
    
7. In the  **More Views** dialog box, click **Organizer**. Scroll through the  **Global (+ non-cached Enterprise)** list to see that **R2 View Test** was automatically added to the global template.
    



