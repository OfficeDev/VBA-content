---
title: "Свойство Document.SurplusShapes (издатель)"
keywords: vbapb10.chm196754
f1_keywords: vbapb10.chm196754
ms.prod: publisher
api_name: Publisher.Document.SurplusShapes
ms.assetid: 8c1c5fee-bea0-1660-a4a5-b465879d6ec9
ms.date: 06/08/2017
ms.openlocfilehash: 8bc7ac17ece4b05cf04b3186bc6e714d0e4b3aaa
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsurplusshapes-property-publisher"></a>Свойство Document.SurplusShapes (издатель)

Возвращает объект **ShapeRange** , представляющий коллекцию избыточные фигур, Microsoft Publisher помещает в разделе **Дополнительное содержимое**в области задач **Формат публикации** после изменения с помощью шаблона документа (мастер) ** [Document.ChangeDocument](document-changedocument-method-publisher.md)** метод или с помощью команды **Изменить шаблон** в пользовательском интерфейсе. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SurplusShapes**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="remarks"></a>Заметки

Publisher классификация фигуры как избыток, если его нельзя отнести новый шаблон после изменения шаблона.


