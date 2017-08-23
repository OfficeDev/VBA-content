---
title: "Объект BorderArts (издатель)"
keywords: vbapb10.chm7798783
f1_keywords: vbapb10.chm7798783
ms.prod: publisher
api_name: Publisher.BorderArts
ms.assetid: 0fc016f6-154e-3591-34b3-e094bbad9d16
ms.date: 06/08/2017
ms.openlocfilehash: 200a0615a89aabf3aeccf24868d32b0cf4c89b2b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderarts-object-publisher"></a>Объект BorderArts (издатель)

Коллекция всех Узорные, доступных для использования в указанной публикации. Узорные — границы предварительно заданных рисунков, которые можно применять для текстовых полей, рамки рисунков или прямоугольники.
 


## <a name="remarks"></a>Заметки

Коллекция **BorderArts** включает все пользовательские типы Узорные, создаваемые пользователем для указанной публикации.
 

 

## <a name="example"></a>Пример

Используйте свойство **[Item](borderarts-item-method-publisher.md)** коллекции **BorderArts** для получения определенного объекта **[Узорные](borderart-object-publisher.md)** . Аргумент Index свойство **Item** может быть номер или имя объекта Узорные.
 

 
В этом примере возвращается Узорные «Apples» из активной публикации. 
 

 



```
Dim bdaTemp As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts.Item (Index:="Apples") 
```

Свойство **[Count](borderarts-count-property-publisher.md)** возвращает число Узорные типы, доступные в указанный документ. Следующий пример показывает число типов Узорные в активном документе.
 

 



```
Sub CountBorderArts() 
 MsgBox ActiveDocument.BorderArts.Count 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](borderarts-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](borderarts-application-property-publisher.md)|
|[Count](borderarts-count-property-publisher.md)|
|[Родительский раздел](borderarts-parent-property-publisher.md)|

