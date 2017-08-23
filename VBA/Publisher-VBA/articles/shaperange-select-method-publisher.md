---
title: "Метод ShapeRange.Select (издатель)"
keywords: vbapb10.chm2293799
f1_keywords: vbapb10.chm2293799
ms.prod: publisher
api_name: Publisher.ShapeRange.Select
ms.assetid: 3252ba74-d051-8c28-a9ed-c6f5ca711dec
ms.date: 06/08/2017
ms.openlocfilehash: 81b2c1156fa1a82fdd905f408b896599214a9ea4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeselect-method-publisher"></a>Метод ShapeRange.Select (издатель)

Выбирает указанный объект.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выберите** ( **_Заменить_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Замена|Необязательный| **Variant**|Указывает, будет ли выделение заменяет предыдущее выделение.  **Значение true** для замены предыдущее выделение с новой выбора; **Значение false** для добавления нового выделения в предыдущем выделение. Значение по умолчанию — **True**.|

## <a name="example"></a>Пример

В этом примере выбирает фигур одним и три по одному в активной публикации.


```vb
ActiveDocument.Pages(1).Shapes.Range(Array(1, 3)).Select
```

В этом примере добавляется фигур двух и четыре по одному в активной публикации в предыдущем выделение.




```vb
ActiveDocument.Pages(1).Shapes.Range(Array(2, 4)) _ 
 .Select Replace:=False
```


