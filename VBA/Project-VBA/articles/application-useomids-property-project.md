---
title: Application.UseOMIDs Property (Project)
keywords: vbapj.chm132781
f1_keywords:
- vbapj.chm132781
ms.prod: project-server
api_name:
- Project.Application.UseOMIDs
ms.assetid: 15339e09-0b65-d939-df47-eb538dee7c38
ms.date: 06/08/2017
---


# Application.UseOMIDs Property (Project)

Gets or sets the corresponding  **Use internal IDs** option on the **Advanced** tab of the **Project Options** dialog box. Read/write **Boolean**.


## Syntax

 _expression_. **UseOMIDs**

 _expression_ A variable that represents an **Application** object.


## Remarks

Object Matching Identifier (OMID) fields are added to objects that can be shared across multilanguage versions. OMIDs are supported for  **Calendar**, **Filter**, **Group**, **Table**, and **View** objects. OMIDs are not supported for **Form** and **Report** objects. Project uses OMIDs to match similar elements in a multilanguage installation and avoid multiple language elements in the UI.


