---
title: "Метод Shape.Select (издатель)"
keywords: vbapb10.chm2228263
f1_keywords: vbapb10.chm2228263
ms.prod: publisher
api_name: Publisher.Shape.Select
ms.assetid: d18914fd-7679-e922-090c-78affdb39d6a
ms.date: 06/08/2017
ms.openlocfilehash: e8dde0717e9861708425773a22fc3fa27b1270c6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeselect-method-publisher"></a>Метод Shape.Select (издатель)

Выбирает указанный объект.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выберите** ( **_Заменить_**)

 переменная _expression_A, представляющий объект **фигуры** .


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


