---
title: "Свойство WizardProperty.Name (издатель)"
keywords: vbapb10.chm1572864
f1_keywords: vbapb10.chm1572864
ms.prod: publisher
api_name: Publisher.WizardProperty.Name
ms.assetid: d66dd4be-9f47-baed-b4aa-6c8cbf293505
ms.date: 06/08/2017
ms.openlocfilehash: 17ffcd54a32f61e73c5576580503bb8bf8d417bd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wizardpropertyname-property-publisher"></a>Свойство WizardProperty.Name (издатель)

Возвращает **строковое** значение, указывающее имя указанного объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляет собой объект- **WizardProperty** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


