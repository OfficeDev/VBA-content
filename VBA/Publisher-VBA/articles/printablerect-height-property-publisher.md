---
title: "Свойство PrintableRect.Height (издатель)"
keywords: vbapb10.chm7536646
f1_keywords: vbapb10.chm7536646
ms.prod: publisher
api_name: Publisher.PrintableRect.Height
ms.assetid: 55d07c00-ee9f-c177-3277-9355618dce6d
ms.date: 06/08/2017
ms.openlocfilehash: 1354bb81a2666294d54a53e837e8ca69abe31b41
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="printablerectheight-property-publisher"></a>Свойство PrintableRect.Height (издатель)

Возвращает значение типа **одного** , представляющий высота в пунктах подготовленных к печати прямоугольника. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Высота**

 переменная _expression_A, представляет собой объект- **PrintableRect** .


## <a name="remarks"></a>Заметки

Допустимые значения для свойства **Height** зависит от размера рабочей области приложения и позиции объекта в рабочей области. По центру объектов на размер страницы не баннер свойство **Height** может быть 0,0-50,0 дюйма. По центру объектов на размер заголовка страницы свойство **Height** может быть 0.0 для 241.0 дюйма.


