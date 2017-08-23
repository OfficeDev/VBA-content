---
title: "Свойство Page.Name (издатель)"
keywords: vbapb10.chm131098
f1_keywords: vbapb10.chm131098
ms.prod: publisher
api_name: Publisher.Page.Name
ms.assetid: cd81994d-506a-69ca-c7f6-472705b2ccd3
ms.date: 06/08/2017
ms.openlocfilehash: d725a01e2a242cadd2cc474e7e3d75e7d4bf35cc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagename-property-publisher"></a>Свойство Page.Name (издатель)

Возвращает или задает **строковое** значение, указывающее имя указанного объекта. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляющий объект **Page** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


