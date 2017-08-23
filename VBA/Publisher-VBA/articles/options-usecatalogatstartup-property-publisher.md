---
title: "Свойство Options.UseCatalogAtStartup (издатель)"
keywords: vbapb10.chm1048612
f1_keywords: vbapb10.chm1048612
ms.prod: publisher
api_name: Publisher.Options.UseCatalogAtStartup
ms.assetid: 7b0cfce9-92f1-5491-c550-421d1c848e0f
ms.date: 06/08/2017
ms.openlocfilehash: 73704f4d3aa4e22b17e6eb5e516acac63c9a630e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsusecatalogatstartup-property-publisher"></a>Свойство Options.UseCatalogAtStartup (издатель)

 **Значение true** для Microsoft Publisher для отображения в каталог при запуске. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UseCatalogAtStartup**

 переменная _expression_A, представляющий объект **параметров** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере задается глобальных параметров для Publisher, включая не отображаются в каталог при запуске.


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


