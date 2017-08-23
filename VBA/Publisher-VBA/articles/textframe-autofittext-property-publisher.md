---
title: "Свойство TextFrame.AutoFitText (издатель)"
keywords: vbapb10.chm3866630
f1_keywords: vbapb10.chm3866630
ms.prod: publisher
api_name: Publisher.TextFrame.AutoFitText
ms.assetid: 468a9d3e-cb9d-8147-60ea-eb839d691e7a
ms.date: 06/08/2017
ms.openlocfilehash: 8b50a265a9714038e629ef318d44e81b9fb885d5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframeautofittext-property-publisher"></a>Свойство TextFrame.AutoFitText (издатель)

Задает или возвращает константу **PbTextAutoFitType**, представляющий как Microsoft Publisher автоматически изменяет размер шрифта текста и размер объектов **TextFrame** для лучшего просмотра. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoFitText**

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

PbTextAutoFitType


## <a name="remarks"></a>Заметки

Значение свойства **AutoFitText** может иметь одно из **[PbTextAutoFitType](pbtextautofittype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Следующий пример проверяет ли надпись содержит текст, и если да, свойству **AutoFitText** лучше всего подходит.


```vb
Sub TextFit() 
 
 Dim tfFrame As TextFrame 
 
 tfFrame = Application.ActiveDocument.MasterPages.Item(1).Shapes(1).TextFrame 
 With tfFrame 
 If .HasText = msoTrue Then .AutoFitText = pbTextAutoFitBestFit 
 End With 
 
End Sub
```


