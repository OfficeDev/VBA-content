---
title: "Свойство Shape.IsExcess (издатель)"
keywords: vbapb10.chm2228377
f1_keywords: vbapb10.chm2228377
ms.prod: publisher
api_name: Publisher.Shape.IsExcess
ms.assetid: 217689d6-7508-92ab-3828-e61fc70f0993
ms.date: 06/08/2017
ms.openlocfilehash: df526847dcce96eafc2ada8ed19b9473a87f95c0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeisexcess-property-publisher"></a>Свойство Shape.IsExcess (издатель)

Указывает, является ли **фигура** родительский объект лишние фигуры после изменения с помощью шаблона документа (мастер) ** [Document.ChangeDocument](document-changedocument-method-publisher.md)** метод или с помощью команды **Изменить шаблон** в пользовательском интерфейсе. Microsoft Publisher помещает все лишние фигуры в разделе **Дополнительное содержимое** в области задач **Формат публикации** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsExcess**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Publisher классификация фигуры как превышение (избыток), если ее нельзя отнести новый шаблон после изменения шаблона.


