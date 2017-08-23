---
title: "Свойство Application.CommandBars (издатель)"
keywords: vbapb10.chm131088
f1_keywords: vbapb10.chm131088
ms.prod: publisher
api_name: Publisher.Application.CommandBars
ms.assetid: 21537c04-d406-6016-4f35-2f6ce6851db2
ms.date: 06/08/2017
ms.openlocfilehash: 41374d7145cba18b97df3f2866d8c93fc9e81918
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationcommandbars-property-publisher"></a>Свойство Application.CommandBars (издатель)

Задает или возвращает коллекцию **CommandBars** , который представляет строки меню и панели инструментов в Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CommandBars**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

CommandBars


## <a name="example"></a>Пример

В этом примере увеличивает все кнопки панели команд, позволяет всплывающие подсказки и отображает все элементы меню при отображении меню.


```vb
Sub CmdBars() 
 
 With CommandBars 
 .LargeButtons = False 
 .DisplayTooltips = True 
 .AdaptiveMenus = False 
 End With 
 
End Sub
```

В этом примере отображаются панели инструментов **объекты** в нижней части окна приложения.




```vb
Sub ShowObjectsToolbar 
 
 With CommandBars("Objects") 
 .Visible = True 
 .Position = msoBarBottom 
 End With 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

