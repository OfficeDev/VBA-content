---
title: Application.ResourceGraphBarStyles Method (Project)
keywords: vbapj.chm2057
f1_keywords:
- vbapj.chm2057
ms.prod: project-server
api_name:
- Project.Application.ResourceGraphBarStyles
ms.assetid: b8d2baf3-7025-e330-a582-451ec0d115c0
ms.date: 06/08/2017
---


# Application.ResourceGraphBarStyles Method (Project)

Sets the styles of bars on the Resource Graph view.


## Syntax

 _expression_. **ResourceGraphBarStyles**( ** _TopLeftShowAs_**, ** _TopLeftColor_**, ** _TopLeftPattern_**, ** _BottomLeftShowAs_**, ** _BottomLeftColor_**, ** _BottomLeftPattern_**, ** _TopRightShowAs_**, ** _TopRightColor_**, ** _TopRightPattern_**, ** _BottomRightShowAs_**, ** _BottomRightColor_**, ** _BottomRightPattern_**, ** _ShowValues_**, ** _ShowAvailabilityLine_**, ** _PercentBarOverlap_**, ** _ProposedLeftShowAs_**, ** _ProposedLeftColor_**, ** _ProposedLeftPattern_**, ** _ProposedRightShowAs_**, ** _ProposedRightColor_**, ** _ProposedRightPattern_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TopLeftShowAs_|Optional|**Integer**|The bar type for the overallocated resources category in the upper-left corner of the  **Bar Styles** dialog box. Can be one of the following **[PjResourceGraphStyle](pjresourcegraphstyle-enumeration-project.md)** constants: **pjBar**, **pjArea**, **pjStep**, **pjLine**, **pjStepLine**, or **pjDoNotShow**.|
| _TopLeftColor_|Optional|**Integer**|The bar color for the overallocated resources category in the upper-left corner of the  **Bar Styles** dialog box. Can be one of the **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _TopLeftPattern_|Optional|**Integer**|The bar pattern for the overallocated resources category in the upper-left corner of the  **Bar Styles** dialog box. Can be one of the **[PjResourceGraphPattern](pjresourcegraphpattern-enumeration-project.md)** constants.|
| _BottomLeftShowAs_|Optional|**Integer**|The bar type for the allocated resources category (the middle left section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _BottomLeftColor_|Optional|**Integer**|The bar color for the allocated resources category (the middle left section) of the  **Bar Styles** dialog box. Can be one of the **PjColor** constants.|
| _BottomLeftPattern_|Optional|**Integer**|The bar pattern for the allocated resources category (the middle left section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _TopRightShowAs_|Optional|**Integer**|The bar type for the overallocated resources category in the upper-right corner of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _TopRightColor_|Optional|**Integer**|The bar color for the overallocated resources category in the upper-right corner of the  **Bar Styles** dialog box. Can be one of the **PjColor** constants.|
| _TopRightPattern_|Optional|**Integer**|The bar pattern for the overallocated resources category in the upper-right corner of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _BottomRightShowAs_|Optional|**Integer**|The bar type for the allocated resources category (the middle right section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _BottomRightColor_|Optional|**Integer**|The bar color for the allocated resources category (the middle right section) of the  **Bar Styles** dialog box. Can be one of the **PjColor** constants.|
| _BottomRightPattern_|Optional|**Integer**|The bar pattern for the allocated resources category (the middle right section) of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _ShowValues_|Optional|**Boolean**|**True** if the corresponding values appear below the bars.|
| _ShowAvailabilityLine_|Optional|**Boolean**|**True** if a horizontal line appears where a resource reaches its maximum availability.|
| _PercentBarOverlap_|Optional|**Integer**|A number from 0 to 100 that specifies the overlap percentage of displayed bars.|
| _ProposedLeftShowAs_|Optional|**Integer**|The bar type for the proposed bookings category in the bottom left section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _ProposedLeftColor_|Optional|**Integer**|The bar color for the proposed bookings category in the bottom left section of the  **Bar Styles** dialog box. Can be one of the **PjColor** constants.|
| _ProposedLeftPattern_|Optional|**Integer**|The bar pattern for the proposed bookings category in the bottom left section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|
| _ProposedRightShowAs_|Optional|**Integer**|The bar type for the proposed bookings category in the bottom right section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphStyle** constants.|
| _ProposedRightColor_|Optional|**Integer**|The bar color for the proposed bookings category in the bottom right section of the  **Bar Styles** dialog box. Can be one of the **PjColor** constants.|
| _ProposedRightPattern_|Optional|**Integer**|The bar pattern for the proposed bookings category in the bottom right section of the  **Bar Styles** dialog box. Can be one of the **PjResourceGraphPattern** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **ResourceGraphBarStyles** method without specifying any arguments displays the **Bar Styles** dialog box.

To edit the resource graph styles where the colors can be specified as hexadecimal RGB values, use the  **[ResourceGraphBarStylesEx](application-resourcegraphbarstylesex-method-project.md)** method


## Example

The following line of code sets proposed resources in the Resource Graph view as a step line in a blue-green color.


```vb
Application.ResourceGraphBarStylesEx ProposedRightShowAs:=pjStepLine, ProposedRightColor:=pjTeal 

```


