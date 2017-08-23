---
title: "Свойство WizardValue.Name (издатель)"
keywords: vbapb10.chm2097152
f1_keywords: vbapb10.chm2097152
ms.prod: publisher
api_name: Publisher.WizardValue.Name
ms.assetid: 51cef04a-e22f-217f-a8a4-d9931057e817
ms.date: 06/08/2017
ms.openlocfilehash: 614de0c966abb989de42634df93b9b7b6bef2932
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardvaluename-property-publisher"></a>Свойство WizardValue.Name (издатель)

Возвращает **строковое** значение, указывающее имя указанного объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляет собой объект- **WizardValue** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


