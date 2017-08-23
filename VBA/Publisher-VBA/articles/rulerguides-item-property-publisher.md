---
title: "Свойство RulerGuides.Item (издатель)"
keywords: vbapb10.chm720896
f1_keywords: vbapb10.chm720896
ms.prod: publisher
api_name: Publisher.RulerGuides.Item
ms.assetid: e0c49279-4fd4-fe61-636c-c29399fdc404
ms.date: 06/08/2017
ms.openlocfilehash: 2339b65fe8881583fb3ebb8c7548fcdaf8eff175
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rulerguidesitem-property-publisher"></a>Свойство RulerGuides.Item (издатель)

Возвращает объект отдельных из указанного семейства сайтов. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **RulerGuides** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Item|Обязательное свойство.| **Длинный**|Количество для возвращаемого объекта.|

## <a name="example"></a>Пример

В этом примере задается положение первого направляющей линейки на 3 дюйма от края публикации.


```vb
ActiveDocument.Pages(1).RulerGuides _ 
 .Item(1).Position = InchesToPoints(3)
```


