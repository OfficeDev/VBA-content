---
title: "Свойство Shape.Name (издатель)"
keywords: vbapb10.chm2228292
f1_keywords: vbapb10.chm2228292
ms.prod: publisher
api_name: Publisher.Shape.Name
ms.assetid: 307c131b-f6ad-38e7-d214-420063d3e5ec
ms.date: 06/08/2017
ms.openlocfilehash: e2d9c3c6eca664f96403bd302b25094edd9bcbeb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapename-property-publisher"></a>Свойство Shape.Name (издатель)

Возвращает или задает **строковое** значение, указывающее имя указанного объекта. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


