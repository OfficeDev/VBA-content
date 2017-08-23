---
title: "Свойство TextStyle.Name (издатель)"
keywords: vbapb10.chm5963782
f1_keywords: vbapb10.chm5963782
ms.prod: publisher
api_name: Publisher.TextStyle.Name
ms.assetid: 54e25e71-83d8-5074-fa0a-f956f075f482
ms.date: 06/08/2017
ms.openlocfilehash: 3e574e4f233b198c71269fdf5743e726fa2118e9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstylename-property-publisher"></a>Свойство TextStyle.Name (издатель)

Возвращает **строковое** значение, указывающее имя указанного объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляющий объект **стиля текста** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


