---
title: "Метод Shapes.AddBuildingBlock (издатель)"
keywords: vbapb10.chm2162768
f1_keywords: vbapb10.chm2162768
ms.prod: publisher
ms.assetid: d875e97e-3519-4a88-916d-ec1a32654581
ms.date: 06/08/2017
ms.openlocfilehash: 8e4ab903c0945740b30fd3dfa17f89c2e5a393f9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddbuildingblock-method-publisher"></a>Метод Shapes.AddBuildingBlock (издатель)

Добавляет объект **[BuildingBlock](buildingblock-object-publisher.md)** и возвращает объект **[фигуры](shape-object-publisher.md)** на странице, который представляет стандартный блок.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddBuildingBlock** ( **_BBlockIn_**, **_слева_** **_сверху_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Тип|Обязательный| **BuildingBlock**|Стандартный блок для возврата как фигуры.|
|Слева|Обязательное свойство.| **Variant**|Положение левого края фигуры, представляющий стандартный блок.|
|Вверх|Обязательное свойство.| **Variant**|Положение верхнего края фигуры, представляющий стандартный блок.|

### <a name="return-value"></a>Возвращаемое значение

 **Фигура**


