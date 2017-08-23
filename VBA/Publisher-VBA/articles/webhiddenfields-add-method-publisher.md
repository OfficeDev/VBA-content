---
title: "Метод WebHiddenFields.Add (издатель)"
keywords: vbapb10.chm3997700
f1_keywords: vbapb10.chm3997700
ms.prod: publisher
api_name: Publisher.WebHiddenFields.Add
ms.assetid: c3035138-f369-b561-b1f8-9977bd9e080c
ms.date: 06/08/2017
ms.openlocfilehash: 38404fab095a8b40be40d22af3e24740db511455
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webhiddenfieldsadd-method-publisher"></a>Метод WebHiddenFields.Add (издатель)

Добавляет нового скрытого поля в веб-формы и возвращает значение типа **Long** , указывающее количество нового поля в коллекции **WebHiddenFields** . Новые поля, всегда помещаются в конце текущий список полей.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Имя_**, **_значение_**)

 переменная _expression_A, представляет собой объект- **WebHiddenFields** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|Обязательное свойство.| **String**|Имя нового поля.|
|Значение|Обязательное свойство.| **String**|Значение нового поля.|

### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

Следующий пример добавляет нового скрытого поля для указанного элемента управления кнопки команды Web. Фигура одно на странице один из активных публикации должен быть элемент управления кнопки команды Web для работы этого примера.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .WebCommandButton.HiddenFields _ 
 .Add Name:="subject", Value:="service request"
```


