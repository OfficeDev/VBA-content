---
title: "Метод PictureFormat.IncrementContrast (издатель)"
keywords: vbapb10.chm3604497
f1_keywords: vbapb10.chm3604497
ms.prod: publisher
api_name: Publisher.PictureFormat.IncrementContrast
ms.assetid: cff50058-2b88-fc2d-633d-411380e5f2f3
ms.date: 06/08/2017
ms.openlocfilehash: 3601c0074a164cbe772b550ab1439dd907a00e78
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatincrementcontrast-method-publisher"></a>Метод PictureFormat.IncrementContrast (издатель)

Изменение контрастности рисунка на указанную величину.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IncrementContrast** ( **_Порядкового номера_**)

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Порядкового номера|Обязательное свойство.| **Один**|Определяет, насколько изменение значение свойства **[контрастности](pictureformat-contrast-property-publisher.md)** рисунка. Положительное значение увеличивает контрастность; отрицательное значение уменьшает контраст. Допустимые значения: от - 1 до 1.|

## <a name="remarks"></a>Заметки

Нельзя настроить контрастность изображения за границу верхней или нижней, для свойства **контрастности** . Например если свойство **контрастности** изначально установлено значение 0,9 и можно указать 0,3 для аргумента **_порядкового номера_** , результирующий уровень контрастности 1.0, являющийся верхний предел для свойства **контрастности** вместо 1.2.

Свойство **контрастности** задать абсолютные яркости для рисунка.


## <a name="example"></a>Пример

В этом примере увеличивает контрастность для всех рисунков на первой странице active публикации, которые уже не задано значение максимального контрастности.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = msoPicture Then 
 shpLoop.PictureFormat.IncrementContrast Increment:=0.1 
 End If 
Next shpLoop 

```


