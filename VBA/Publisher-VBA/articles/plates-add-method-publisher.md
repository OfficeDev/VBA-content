---
title: "Метод Plates.Add (издатель)"
keywords: vbapb10.chm2818052
f1_keywords: vbapb10.chm2818052
ms.prod: publisher
api_name: Publisher.Plates.Add
ms.assetid: 7fb7b602-8797-e275-4ff7-2e87cf1db11f
ms.date: 06/08/2017
ms.openlocfilehash: e7bc79f2d33269c847259a6900ae51e93d83450f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="platesadd-method-publisher"></a>Метод Plates.Add (издатель)

Добавляет новый цвет формы на указанный объект **формы** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_PlateColor_**)

 переменная _expression_A, представляющий объект **формы** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PlateColor|Необязательный| **ColorFormat**| Объект **ColorFormat** . Параметры цвета, применяемые к новой формы.|

## <a name="remarks"></a>Заметки

Если ** [ColorMode](http://msdn.microsoft.com/library/58befa97-9d9b-9294-18b2-ae10dc87f51c%28Office.15%29.aspx)** свойства указанного публикации не **pbColorModeSpot** или **pbColorModeSpotAndProcess**, возникает ошибка.


## <a name="example"></a>Пример

Следующий пример добавляет цвет формы active публикации при публикации.


```vb
If ActiveDocument.ColorMode = pbColorModeSpot Then 
 ActiveDocument.Plates.Add 
End If
```


