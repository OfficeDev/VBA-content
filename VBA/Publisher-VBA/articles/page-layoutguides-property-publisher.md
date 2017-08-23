---
title: "Свойство Page.LayoutGuides (издатель)"
keywords: vbapb10.chm393270
f1_keywords: vbapb10.chm393270
ms.prod: publisher
api_name: Publisher.Page.LayoutGuides
ms.assetid: eb9ac463-2b9f-9c68-b58f-6d93fe4993c8
ms.date: 06/08/2017
ms.openlocfilehash: b8ec4be34258f70b474278975f2dc6ea9fe2e121
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagelayoutguides-property-publisher"></a>Свойство Page.LayoutGuides (издатель)

Возвращает объект **[LayoutGuides](layoutguides-object-publisher.md)** , состоящий из полей и сетки направляющие разметки для всех страниц, включая главные страницы в публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LayoutGuides**

 переменная _expression_A, представляющий объект **Page** .


## <a name="example"></a>Пример

В следующем примере изменяется направляющие сетки, чтобы существует три столбца и пять строк.


```vb
Dim layTemp As LayoutGuides 
 
Set layTemp = ActiveDocument.LayoutGuides 
 
With layTemp 
 .Rows = 5 
 .Columns = 3 
End With 

```


