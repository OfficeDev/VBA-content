---
title: "Свойство Tag.Name (издатель)"
keywords: vbapb10.chm4718595
f1_keywords: vbapb10.chm4718595
ms.prod: publisher
api_name: Publisher.Tag.Name
ms.assetid: a35e8c51-e4c8-2554-eb44-8f202795fbc7
ms.date: 06/08/2017
ms.openlocfilehash: b63d6614ce76b94fd187249564a44c4c05e168b0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tagname-property-publisher"></a>Свойство Tag.Name (издатель)

Возвращает **строковое** значение, указывающее имя указанного объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляет собой объект- **тег** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


