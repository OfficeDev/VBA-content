---
title: Page Object
keywords: fm20.chm2000590
f1_keywords:
- fm20.chm2000590
ms.prod: office
api_name:
- Office.Page
ms.assetid: 889faad0-d2ce-b404-a603-2a491c27df23
ms.date: 06/08/2017
---


# Page Object



One page of a  **MultiPage** and a single member of a **Pages** collection.
 **Remarks**
Each  **Page** object contains its own set of controls and does not necessarily rely on other pages in the[collection](vbe-glossary.md) for information. A **Page** inherits some properties from its[container](vbe-glossary.md); the value of each [inherited property](glossary-vba.md) is set by the container.
A  **Page** has a unique name and index value within a **Pages** collection. You can reference a **Page** by either its name or its index value. The index of the first **Page** in a collection is 0; the index of the second **Page** is 1; and so on. When two **Page** objects have the same name, you must reference each **Page** by its index value. References to the name in code will access only the first **Page** that uses the name.
The default name for the first  **Page** is Page1; the default name for the second **Page** is Page2.

## Related Topics

[ Page Object](http://msdn.microsoft.com/library/2669ab2b-1dfc-47ef-bcaf-e8a9773f010b%28Office.15%29.aspx)


