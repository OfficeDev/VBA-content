---
title: "Свойство Document.AdvancedPrintOptions (издатель)"
keywords: vbapb10.chm196713
f1_keywords: vbapb10.chm196713
ms.prod: publisher
api_name: Publisher.Document.AdvancedPrintOptions
ms.assetid: 33c075e0-f813-9bb4-e199-96e5e9ed4ba8
ms.date: 06/08/2017
ms.openlocfilehash: acfb6ecfb03557d5486ca6d79e3e6b0530ea4772
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentadvancedprintoptions-property-publisher"></a>Свойство Document.AdvancedPrintOptions (издатель)

Возвращает объект, представляющий параметры печати для публикации на **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AdvancedPrintOptions**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

AdvancedPrintOptions


## <a name="remarks"></a>Заметки

Свойства объекта **AdvancedPrintOptions** соответствуют параметрам в диалоговом окне **Дополнительные параметры печати** .


## <a name="example"></a>Пример

Следующий пример проверяет, чтобы определить, установлено ли active публикации для печати цветоделение. Если Да, оно установлено для печати форм только для красок, используемые в публикации, а также не печатать формы для всех страниц, где не используется цвет.


```vb
Sub PrintOnlyInksUsed 
 With ActiveDocument.AdvancedPrintOptions 
 If .PrintMode = pbPrintModeSeparations Then 
 .InksToPrint = pbInksToPrintUsed 
 .PrintBlankPlates = False 
 End If 
 End With 
End Sub
```


