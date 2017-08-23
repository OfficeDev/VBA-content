---
title: "Свойство Document.ScratchArea (издатель)"
keywords: vbapb10.chm196657
f1_keywords: vbapb10.chm196657
ms.prod: publisher
api_name: Publisher.Document.ScratchArea
ms.assetid: 782d9b7f-b620-60f0-c21d-04f588c37cc6
ms.date: 06/08/2017
ms.openlocfilehash: 0af19ac3f646473c90e934fac077ff1199e95c80
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentscratcharea-property-publisher"></a>Свойство Document.ScratchArea (издатель)

Возвращает объект **[ScratchArea](scratcharea-object-publisher.md)** для данного документа.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ScratchArea**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

ScratchArea


## <a name="remarks"></a>Заметки

Объект **ScratchArea** представляет собой коллекцию объектов на странице "Рабочий". Объект **ScratchArea** является не в коллекции **страниц** , так как он не существенно страницы; его только сходство на страницу — это, что он может содержать объекты.


## <a name="example"></a>Пример

В этом примере задаются объект переменной как первую фигуру вспомогательной области активных документов.


```vb
Sub ScratchPad() 
 
 Dim saPage As ScratchArea 
 Dim objFirst As Object 
 
 saPage = Application.ActiveDocument.ScratchArea 
 objFirst = saPage.Shapes(1) 
 
End Sub
```


