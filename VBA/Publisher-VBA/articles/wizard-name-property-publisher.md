---
title: "Свойство Wizard.Name (издатель)"
keywords: vbapb10.chm1441796
f1_keywords: vbapb10.chm1441796
ms.prod: publisher
api_name: Publisher.Wizard.Name
ms.assetid: 1e0a7ec6-87ee-7c26-cf98-e849c5617e58
ms.date: 06/08/2017
ms.openlocfilehash: b6a0c483bd1e7607aa4634fd0f78c263f74be67b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardname-property-publisher"></a>Свойство Wizard.Name (издатель)

Возвращает **строковое** значение, указывающее имя указанного объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляющий объект **мастера** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


