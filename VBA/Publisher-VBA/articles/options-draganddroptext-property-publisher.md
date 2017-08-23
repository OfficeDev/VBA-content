---
title: "Свойство Options.DragAndDropText (издатель)"
keywords: vbapb10.chm1048584
f1_keywords: vbapb10.chm1048584
ms.prod: publisher
api_name: Publisher.Options.DragAndDropText
ms.assetid: 55fb68e8-4ddc-6866-00d8-bdd6a1e25ec3
ms.date: 06/08/2017
ms.openlocfilehash: df6d8ebc32dbeeee4135045495872e089fa320f3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsdraganddroptext-property-publisher"></a>Свойство Options.DragAndDropText (издатель)

 **Значение true,** чтобы разрешить перетаскивание текста. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DragAndDropText**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере задается глобальных параметров для Microsoft Publisher, включая активацию перетаскивания для изменения положения текста.


```vb
Sub SetGlobalOptions() 
 With Options 
 .AutoFormatWord = True 
 .AutoKeyboardSwitching = True 
 .AutoSelectWord = True 
 .DragAndDropText = True 
 .UseCatalogAtStartup = False 
 .UseHelpfulMousePointers = False 
 End With 
End Sub
```


