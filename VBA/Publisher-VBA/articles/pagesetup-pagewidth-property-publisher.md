---
title: "Свойство PageSetup.PageWidth (издатель)"
keywords: vbapb10.chm6946822
f1_keywords: vbapb10.chm6946822
ms.prod: publisher
api_name: Publisher.PageSetup.PageWidth
ms.assetid: 547f2881-d9fa-fa5f-1643-ab08084fb423
ms.date: 06/08/2017
ms.openlocfilehash: e069aabfb4569fe1062d4b48b229726050a6fc5c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetuppagewidth-property-publisher"></a>Свойство PageSetup.PageWidth (издатель)

Возвращает или задает **Variant** , который представляет ширину страниц в публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageWidth**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются как точки. Строковые значения можно в любое устройство, поддерживаемый Microsoft Publisher (например, «2,5 дюйма»). Допустимый диапазон допустимых значений — от 0 до различие между ширину листа и ширину страницы.


## <a name="example"></a>Пример

В следующем примере задается в ширину восемь дюйма для страниц в активной публикации.


```vb
Public Sub PageWidth_Example() 
 ActiveDocument.PageSetup.PageWidth = InchesToPoints(8) 
End Sub
```


