---
title: Application.ShareProjectOnline Method (Project)
keywords: vbapj.chm74
f1_keywords:
- vbapj.chm74
ms.prod: project-server
api_name:
- Project.Application.ShareProjectOnline
ms.assetid: 7742715a-d78a-334b-5655-7047efd28890
ms.date: 06/08/2017
---


# Application.ShareProjectOnline Method (Project)

Opens the URL for information about sharing projects in the  **Share with Project Online** section in the Backstage view.


## Syntax

 _expression_. **ShareProjectOnline**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **ShareProjectOnline** method opens the same URL that you see when you choose **Learn More** in the **Share with Project Online** section in the Backstage view. The URL is `http://office.microsoft.com/projectserver/`.


 **Note**  The  **Share with Project Online** section in the Backstage view is visible only when the **Online** value exists as a **DWord** value = 1, in the `HKCU\Software\Microsoft\Office\15.0\MS Project\Options\General` key of the Windows registry. When the **Online** value = 0, the **Share with Project Online** section is hidden.


