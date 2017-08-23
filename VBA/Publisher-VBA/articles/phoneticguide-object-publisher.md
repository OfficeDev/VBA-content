---
title: "Объект PhoneticGuide (издатель)"
keywords: vbapb10.chm6225919
f1_keywords: vbapb10.chm6225919
ms.prod: publisher
api_name: Publisher.PhoneticGuide
ms.assetid: 164e8b54-4bad-4de9-bf6e-52c5687dfbc6
ms.date: 06/08/2017
ms.openlocfilehash: 1f56cadf408dbffc535626c4ed03e4fc59455023
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="phoneticguide-object-publisher"></a>Объект PhoneticGuide (издатель)

Представляет базовый текста с помощью дополнительных текст, отображаемый над текстом в соответствии с транскрипцию.
 


## <a name="example"></a>Пример

Свойство **PhoneticGuide** объекта **поля** для возврата существующего объекта **PhoneticGuide** . Метод **AddPhoneticGuide** коллекции **полей** для создания нового объекта **PhoneticGuide** .
 

 

 

 
Следующий пример добавляет новый объект **PhoneticGuide** active публикации.
 

 



```
Selection.TextRange.Fields.AddPhoneticGuide _ 
 Range:=Selection.TextRange, Text:="ver-E nIs", _ 
 Alignment:=pbPhoneticGuideAlignmentCenter, _ 
 Raise:=11, FontSize:=7
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Очистить](phoneticguide-clear-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Выравнивание](phoneticguide-alignment-property-publisher.md)|
|[Приложения](phoneticguide-application-property-publisher.md)|
|[BaseText](phoneticguide-basetext-property-publisher.md)|
|[FontName](phoneticguide-fontname-property-publisher.md)|
|[FontSize](phoneticguide-fontsize-property-publisher.md)|
|[Родительский раздел](phoneticguide-parent-property-publisher.md)|
|[Чтобы увеличить](phoneticguide-raise-property-publisher.md)|
|Да|

