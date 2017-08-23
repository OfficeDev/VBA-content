---
title: "Свойство Document.Name (издатель)"
keywords: vbapb10.chm196630
f1_keywords: vbapb10.chm196630
ms.prod: publisher
api_name: Publisher.Document.Name
ms.assetid: fcf86fcc-a3aa-b4c6-1ecc-202972ac558b
ms.date: 06/08/2017
ms.openlocfilehash: 50ebad88d9554e364e9112a271183799ede86a8f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentname-property-publisher"></a>Свойство Document.Name (издатель)

Возвращает **строковое** значение, указывающее имя указанного объекта. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

Имя объекта можно использовать в сочетании с **элемента** метод или свойство **Item** возвращает ссылку на объект, если **элемент** метод или свойство для семейства сайтов, содержащее объект принимает аргумент **типа Variant** . Например, если значение свойства **Name** для фигуры — 2 прямоугольника, затем `.Shapes("Rectangle 2")` возвращает ссылку на фигуры.

Свойство **Name** является свойством по умолчанию для объектов **Узорные**, **BorderArtFormat**и **метки** .


