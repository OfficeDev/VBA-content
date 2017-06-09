---
title: Application.ResourceSharing Method (Project)
keywords: vbapj.chm105
f1_keywords:
- vbapj.chm105
ms.prod: project-server
api_name:
- Project.Application.ResourceSharing
ms.assetid: c11f9715-83c2-7872-1d53-fb538ed21c74
ms.date: 06/08/2017
---


# Application.ResourceSharing Method (Project)

Controls resource sharing, for local resources and projects.


## Syntax

 _expression_. **ResourceSharing**( ** _Share_**, ** _Name_**, ** _Pool_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Share_|Optional|**Boolean**|**True** if local resources are shared. If **Name** is specified, **Share** is ignored.|
| _Name_|Optional|**String**|The file name of the project that contains the local resource pool.|
| _Pool_|Optional|**Boolean**|**True** if resources in the local pool take precedence over resources in the project .|

### Return Value

 **Boolean**


## Remarks

Using the  **ResourceSharing** method without specifying any arguments displays the **ShareResources** dialog box.


 **Note**  Project Professional can share local resources only when not logged on Project Server. If Project Professional is using a Project Server profile, local resource sharing is unavailable.


## Example

In the following example, the project that contains the resources to share is named SharedResourcePool.mpp. If the active project is named Sharer.mpp, the code enables Sharer.mpp to access resources from SharedResourcePool.mpp, where resources in the pool take precedence. Both projects must be open.


```vb
Application.ResourceSharing Share:=False, Name:="SharedResourcePool.mpp", Pool:=True
```


