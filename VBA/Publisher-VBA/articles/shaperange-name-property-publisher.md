---
title: "Свойство ShapeRange.Name (издатель)"
keywords: vbapb10.chm2293828
f1_keywords: vbapb10.chm2293828
ms.prod: publisher
api_name: Publisher.ShapeRange.Name
ms.assetid: 517eca4b-fa8c-0f6a-2829-75704bb4c899
ms.date: 06/08/2017
ms.openlocfilehash: 23473bf9d7a5bdd135d38977778c6dae38e54a94
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangename-property-publisher"></a>Свойство ShapeRange.Name (издатель)

Возвращает или задает **строковое** значение, указывающее имя указанного объекта. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


