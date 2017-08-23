---
title: "Свойство CellRange.Height (издатель)"
keywords: vbapb10.chm5177348
f1_keywords: vbapb10.chm5177348
ms.prod: publisher
api_name: Publisher.CellRange.Height
ms.assetid: c4367357-5c33-7461-0cb4-a3fc392bc4bc
ms.date: 06/08/2017
ms.openlocfilehash: 55a30217a38d425a6bf6a6279ec113d8c40aaeff
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrangeheight-property-publisher"></a>Свойство CellRange.Height (издатель)

Возвращает значение типа **Long** , представляющих высота (в ячейках) таблицы, диапазон ячеек или страницы. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Высота**

 переменная _expression_A, представляет собой объект- **CellRange** .


## <a name="remarks"></a>Заметки

Допустимые значения для свойства **Height** зависит от размера рабочей области приложения и позиции объекта в рабочей области. По центру объектов на размер страницы не баннер свойство **Height** может быть 0,0-50,0 ячеек. По центру объектов на размер заголовка страницы свойство **Height** может быть 0.0 для 241.0 ячеек.


