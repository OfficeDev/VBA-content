---
title: "Свойство Options.AddHebDoubleQuote (издатель)"
keywords: vbapb10.chm1048629
f1_keywords: vbapb10.chm1048629
ms.prod: publisher
api_name: Publisher.Options.AddHebDoubleQuote
ms.assetid: 9c71b52e-0273-7ca9-1f50-5beed65c2e73
ms.date: 06/08/2017
ms.openlocfilehash: 4ea28b58e4e56a5937e8149387df643aa5565a66
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsaddhebdoublequote-property-publisher"></a>Свойство Options.AddHebDoubleQuote (издатель)

 **Значение true** для Microsoft Publisher для отображения двойные кавычки для Нумерация буквами иврита. Значение по умолчанию — **False**. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddHebDoubleQuote**

 переменная _expression_A, представляющий объект **параметров** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство доступно только в том случае, если иврит включена для Microsoft Office на вашем компьютере. 

Это свойство применяется только к Нумерация буквами иврита.

Как все свойства объекта **[Options](options-object-publisher.md)** , текущее значение свойства **AddHebDoubleQuote** становится значение по умолчанию, применяемые к все новые публикации.

Это свойство соответствует флажок **Добавить двойные кавычки для иврита алфавита нумерации** в диалоговом окне **список** .


## <a name="example"></a>Пример

В следующем примере задается Publisher для отображения двойные кавычки для Нумерация буквами иврита.


```vb
Publisher.Options.AddHebDoubleQuote = True
```


