---
title: ContactCard.Show Method (Office)
ms.prod: office
api_name:
- Office.ContactCard.Show
ms.assetid: 57fe503a-3298-0bec-3c26-31ae88aa6534
ms.date: 06/08/2017
---


# ContactCard.Show Method (Office)

Displays the contact card at the specified x-coordinate position outside the specified rectangle. 


## Syntax

 _expression_. **Show**( **_Style_**, **_Left_**, **_Right_**, **_Top_**, **_Bottom_**, **_xcord_**, **_fDelay_** )

 _expression_ An expression that returns a **ContactCard** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**MsoContactCardStyle**|Determines whether the card is displayed as a hover card or as a fully expanded card. See Remarks for possible values.|
| _Left_|Required|**Long**|Specifies the x-coordinate of the left side of the rectangle where the card is not displayed.|
| _Right_|Required|**Long**|Specifies the x-coordinate of the right side of the rectangle where the card is not displayed.|
| _Top_|Required|**Long**|Specifies the y-coordinate of the top side of the rectangle where the card is not displayed.|
| _Bottom_|Required|**Long**|Specifies the y-coordinate of the bottom side of the rectangle where the card is not displayed.|
| _xcord_|Required|**Long**|Specifies the x-coordinate position of the left edge of the card.|
| _fDelay_|Required|**Boolean**|Determines if there is a delay before the card is displayed. |

### Return Value

Nothing


## Remarks

 _Style_ must be one of the following **MsoContactCardStyle** values.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**msoContactCardTypeEnterpriseContact**|0|Represents a contact card for an enterprise contact address.|
|**msoContactCardTypePersonalContact**|1|Represents a contact card for a personal contact address.|
|**msoContactCardTypeUnknownContact**|2|Represents a contact card for an unknown contact address.|
|**msoContactCardTypeEnterpriseGroup**|3|Represents a contact card for an enterprise distribution list contact address.|
|**msoContactCardTypePersonalDistributionList**|4|Represents a contact card for a personal distribution list contact address.|
 _fDelay_ applies only if _Style_ is **msoContactCardStyleHover**. For all other card styles, _fDelay_ is ignored.


## See also


#### Concepts


[ContactCard Object](contactcard-object-office.md)
#### Other resources


[ContactCard Object Members](contactcard-members-office.md)

