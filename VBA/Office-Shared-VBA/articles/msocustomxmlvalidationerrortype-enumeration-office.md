---
title: MsoCustomXMLValidationErrorType Enumeration (Office)
ms.prod: office
api_name:
- Office.MsoCustomXMLValidationErrorType
ms.assetid: db2acb55-ce1b-8b2e-1539-45c63f39f557
ms.date: 06/08/2017
---


# MsoCustomXMLValidationErrorType Enumeration (Office)

Indicates how validation errors will be cleared or generated.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**msoCustomXMLValidationErrorAutomaticallyCleared**|1|Specifies that the error will clear itself whenever any change is made to the node it is bound to. |
|**msoCustomXMLValidationErrorManual**|2|Specifies that the error will not be cleared until the  **Delete** method is called.|
|**msoCustomXMLValidationErrorSchemaGenerated**|0|Specifies that where there is a non-empty schema collection available for the custom XML part and validation is in effect, any changes to the part will cause validation errors.|

