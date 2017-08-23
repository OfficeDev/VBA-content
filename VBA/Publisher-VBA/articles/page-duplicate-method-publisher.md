---
title: "Метод Page.Duplicate (издатель)"
keywords: vbapb10.chm393256
f1_keywords: vbapb10.chm393256
ms.prod: publisher
api_name: Publisher.Page.Duplicate
ms.assetid: 9ef9d493-d2ca-8cac-3cce-6f0878acb288
ms.date: 06/08/2017
ms.openlocfilehash: 9758821cd621c9fb1bfbe4200adc8b6ca3205d4e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageduplicate-method-publisher"></a>Метод Page.Duplicate (издатель)

Создает копию на указанный объект **страницы** и возвращает новый объект **Page** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Дублирующиеся**

 переменная _expression_A, представляющий объект **Page** .


### <a name="return-value"></a>Возвращаемое значение

Page


## <a name="example"></a>Пример

В следующем примере дублирует первой страницы публикации и затем задает свойства для повторяющихся. Фигура добавляется новая страница и установки свойств фигуры.


```vb
Dim objPage As Page 
Set objPage = ActiveDocument.Pages(1).Duplicate 
With objPage 
 .Background.Fill.ForeColor.SchemeColor = pbSchemeColorAccent1 
 .Shapes.AddShape msoShapeRectangle, 150, 250, 310, 275 
 With .Shapes(1) 
 .Fill.ForeColor.SchemeColor = pbSchemeColorAccent3 
 End With 
End With 

```


