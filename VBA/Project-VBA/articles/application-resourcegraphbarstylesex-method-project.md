---
title: Application.ResourceGraphBarStylesEx Method (Project)
keywords: vbapj.chm2153
f1_keywords:
- vbapj.chm2153
ms.prod: project-server
api_name:
- Project.Application.ResourceGraphBarStylesEx
ms.assetid: 903c3894-77c9-bd0a-dee0-85c7fcadea38
ms.date: 06/08/2017
---


# Application.ResourceGraphBarStylesEx Method (Project)

Sets the styles of bars on the Resource Graph view, where colors can be hexadecimal values. 


## Syntax

 _expression_. **ResourceGraphBarStylesEx**( ** _TopLeftShowAs_**, ** _TopLeftColor_**, ** _TopLeftPattern_**, ** _BottomLeftShowAs_**, ** _BottomLeftColor_**, ** _BottomLeftPattern_**, ** _TopRightShowAs_**, ** _TopRightColor_**, ** _TopRightPattern_**, ** _BottomRightShowAs_**, ** _BottomRightColor_**, ** _BottomRightPattern_**, ** _ShowValues_**, ** _ShowAvailabilityLine_**, ** _PercentBarOverlap_**, ** _ProposedLeftShowAs_**, ** _ProposedLeftColor_**, ** _ProposedLeftPattern_**, ** _ProposedRightShowAs_**, ** _ProposedRightColor_**, ** _ProposedRightPattern_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TopLeftShowAs_|Optional|**Integer**|The bar type for the overallocated resources category in the upper-left corner of the  **Bar Styles** dialog box. Can be one of the following **[PjResourceGraphStyle](pjresourcegraphstyle-enumeration-project.md)** constants: **pjBar**, **pjArea**, **pjStep**, **pjLine**, **pjStepLine**, or **pjDoNotShow**.|
| _TopLeftColor_|Optional|**Integer**|The bar color for the overallocated resources category in the upper-left corner of the  **Bar Styles** dialog box. Can be a hexadecimal value, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|
| _TopLeftPattern_|Optional|**Integer**|The bar pattern for the overallocated resources category in the upper-left corner of the  **Bar Styles** dialog box. Can be one of the **[PjResourceGraphPattern](pjresourcegraphpattern-enumeration-project.md)** constants.|
| _BottomLeftShowAs_|Optional|**Integer**|The bar type for the allocated resources category (the middle left section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _BottomLeftColor_|Optional|**Integer**|The bar color for the allocated resources category (the middle left section) of the  **Bar Styles** dialog box. Can be a hexadecimal value, where red is the last byte. For example, the value &;HFF00 is green.|
| _BottomLeftPattern_|Optional|**Integer**|The bar pattern for the allocated resources category (the middle left section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _TopRightShowAs_|Optional|**Integer**|The bar type for the overallocated resources category in the upper-right corner of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _TopRightColor_|Optional|**Integer**|The bar color for the overallocated resources category in the upper-right corner of the  **Bar Styles** dialog box. Can be a hexadecimal value, where red is the last byte.|
| _TopRightPattern_|Optional|**Integer**|The bar pattern for the overallocated resources category in the upper-right corner of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _BottomRightShowAs_|Optional|**Integer**|The bar type for the allocated resources category (the middle right section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _BottomRightColor_|Optional|**Integer**|The bar color for the allocated resources category (the middle right section) of the  **Bar Styles** dialog box. Can be a hexadecimal value, where red is the last byte.|
| _BottomRightPattern_|Optional|**Integer**|The bar pattern for the allocated resources category (the middle right section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _ShowValues_|Optional|**Boolean**|**True** if the corresponding values appear below the bars.|
| _ShowAvailabilityLine_|Optional|**Boolean**|**True** if a horizontal line appears where a resource reaches its maximum availability.|
| _PercentBarOverlap_|Optional|**Integer**|A number from 0 to 100 that specifies the overlap percentage of displayed bars.|
| _ProposedLeftShowAs_|Optional|**Integer**|The bar type for the proposed bookings category in the bottom left section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _ProposedLeftColor_|Optional|**Integer**|The bar color for the proposed bookings category in the bottom left section of the  **Bar Styles** dialog box. Can be a hexadecimal value, where red is the last byte.|
| _ProposedLeftPattern_|Optional|**Integer**|The bar pattern for the proposed bookings category in the bottom left section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _ProposedRightShowAs_|Optional|**Integer**|The bar type for the proposed bookings category in the bottom right section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _ProposedRightColor_|Optional|**Integer**|The bar color for the proposed bookings category in the bottom right section of the  **Bar Styles** dialog box. Can be a hexadecimal value, where red is the last byte.|
| _ProposedRightPattern_|Optional|**Integer**|The bar pattern for the proposed bookings category in the bottom right section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **ResourceGraphBarStylesEx** method without specifying any arguments displays the **Bar Styles** dialog box.




 **Note**  If you use any of the  **PjColor** enumeration constants for the color parameters, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the ResourceGraphBarStylesEx method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[ResourceGraphBarStyles](application-resourcegraphbarstyles-method-project.md)** method.


## Example

The following line of code sets proposed resources in the Resource Graph view as a step line in a light blue-green color.


```vb
Application.ResourceGraphBarStylesEx ProposedRightShowAs:=pjStepLine, ProposedRightColor:=&;HD0FF00 

```


