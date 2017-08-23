---
title: "Метод Columns.Item (издатель)"
keywords: vbapb10.chm5046272
f1_keywords: vbapb10.chm5046272
ms.prod: publisher
api_name: Publisher.Columns.Item
ms.assetid: c16df25c-ea8d-c04e-bccd-7e642bb7198a
ms.date: 06/08/2017
ms.openlocfilehash: 84f39d0f40ea35ca452944577abd032651f6af20
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="columnsitem-method-publisher"></a>Метод Columns.Item (издатель)

Возвращает объект отдельных **столбцов** в указанной коллекции **столбцов** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **столбцов** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Количество для возвращаемого объекта.|

### <a name="return-value"></a>Возвращаемое значение

Столбец


## <a name="example"></a>Пример

В этом примере возвращается первый столбец из коллекции **столбцов** .


```vb
Dim colTemp As Column 
 
Set colTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).Table.Columns.Item(Index:=1)
```


