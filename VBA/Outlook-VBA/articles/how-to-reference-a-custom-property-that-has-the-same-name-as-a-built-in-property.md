---
title: "How to: Reference a Custom Property that Has the Same Name as a Built-in Property of the Control"
keywords: olfm10.chm3077225
f1_keywords:
- olfm10.chm3077225
ms.prod: outlook
ms.assetid: 55b2f832-6c23-c43d-0253-1b73f745e1b6
ms.date: 06/08/2017
---


# How to: Reference a Custom Property that Has the Same Name as a Built-in Property of the Control

Assume a new control has a  **Top** property that is different from the standard **Top** property in Microsoft Forms. You can use either property, based on the syntax:


- 
```
  control.Top
```


    uses the standard  **Top** property.
    
- 
```
  control.Object.Top
```


    uses the  **Top** property from the added control.
    

