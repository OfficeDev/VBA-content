---
title: "Свойство Application.OfficeDataSourceObject (издатель)"
keywords: vbapb10.chm131123
f1_keywords: vbapb10.chm131123
ms.prod: publisher
api_name: Publisher.Application.OfficeDataSourceObject
ms.assetid: d7262328-d5b6-6f55-d8c1-e6c072e29e3f
ms.date: 06/08/2017
ms.openlocfilehash: 1d68403770d0b1386b55e3047b18fae2b5b686cc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationofficedatasourceobject-property-publisher"></a>Свойство Application.OfficeDataSourceObject (издатель)

Возвращает объект **OfficeDataSourceObject** , представляющий источника данных в операции объединения слияния почты и каталогов. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OfficeDataSourceObject**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

OfficeDataSourceObject


## <a name="example"></a>Пример

Следующий пример отображает сведения о текущей источника данных для слияния почты.


```vb
Dim odsoTemp As Office.OfficeDataSourceObject 
 
Set odsoTemp = Application.OfficeDataSourceObject 
 
With odsoTemp 
 Debug.Print "Connection string: " &; .ConnectString 
 Debug.Print "Data source: " &; .DataSource 
 Debug.Print "Table: " &; .Table 
End With
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

