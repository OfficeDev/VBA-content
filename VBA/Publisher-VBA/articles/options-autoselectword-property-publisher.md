---
title: "Свойство Options.AutoSelectWord (издатель)"
keywords: vbapb10.chm1048581
f1_keywords: vbapb10.chm1048581
ms.prod: publisher
api_name: Publisher.Options.AutoSelectWord
ms.assetid: 2b36f0d2-3260-aa3d-13b2-ae08b8d631d1
ms.date: 06/08/2017
ms.openlocfilehash: 50d097ba74f6b11be90983159c12e1aa7f687b19
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsautoselectword-property-publisher"></a>Свойство Options.AutoSelectWord (издатель)

 **Значение true** для Microsoft Publisher для автоматического выбора слово целиком при выделении текста. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoSelectWord**

 переменная _expression_A, представляющий объект **параметров** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере задается Publisher глобальные параметры, включая активацию автоматически Выбор целого слова при выделении текста.


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


