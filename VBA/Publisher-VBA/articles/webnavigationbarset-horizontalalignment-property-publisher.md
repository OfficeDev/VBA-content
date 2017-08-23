---
title: "Свойство WebNavigationBarSet.HorizontalAlignment (издатель)"
keywords: vbapb10.chm8519688
f1_keywords: vbapb10.chm8519688
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.HorizontalAlignment
ms.assetid: 7d615a5a-793c-fd78-3dca-a268740b67aa
ms.date: 06/08/2017
ms.openlocfilehash: b584e11daf5b684f6bc9d8036416938031afcecc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsethorizontalalignment-property-publisher"></a>Свойство WebNavigationBarSet.HorizontalAlignment (издатель)

Задает или возвращает константу **PbWizardNavBarAlignment** , который представляет горизонтальное выравнивание кнопок в наборе панель навигации Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalAlignment**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

PbWizardNavBarAlignment


## <a name="remarks"></a>Заметки

Это свойство используется для задания способ отображения кнопок в наборе панель навигации горизонтальный Web. Например объект **WebNavigationBarSet** , содержащий 5 ссылок с помощью свойства **HorizontalButtonCount** значение 3, а свойство **HorizontalAlignment** , задайте значение **pbnbAlignRight** выравнивания кнопок в таблице 1 строкой и 3 столбцов. В первой строке будет сначала 3 кнопки и оставшиеся 2 кнопки будут добавлены в правые столбцы второй строке.

Возвращает значение «Доступ запрещен», если **IsHorizontal** = **значение False** для указанного объекта **WebNavigationBarSet** . Чтобы установить ориентацию панель навигации, равной горизонтальной первоначального перед установкой **HorizontalAlignment** , используйте метод **ChangeOrientation** .

Значение свойства **HorizontalAlignment** может быть присвоено любое из **[PbWizardNavBarAlignment](pbwizardnavbaralignment-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В следующем примере возвращается первый панель навигации с активного документа изменения ориентации на горизонтальную при необходимости задает для свойства **HorizontalButtonCount** значение 3 и свойству **HorizontalAlignment** **pbnbAlignRight**.


```vb
With ActiveDocument.WebNavigationBarSets(1) 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignRight 
End With
```


