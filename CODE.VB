Function GET_JITA_BUY(type_id)
Dim strJson As String
Dim xmlHttp As Object
Dim index_last As Long
Dim index_write As Long
Dim price As Double
Dim lowprice As Double
station_id = "60003760"
Url = "https://esi.evetech.net/latest/markets/10000002/orders/?datasource=tranquility&order_type=sell&type_id=" + CStr(type_id)
Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
xmlHttp.Open "GET", Url, False
xmlHttp.send
If xmlHttp.ReadyState = 4 Then
strJson = xmlHttp.responseText
End If
index_last = InStrRev(strJson, "location_id" & Chr(34) & ":" & station_id)
index_write = 1
highprice = 0
Do While index_write < index_last
retprice = GetPrice(strJson, index_write)
price = retprice
If price > highprice Then
highprice = price
End If
index_write = InStr(index_write + 1, strJson, "location_id" & Chr(34) & ":" & station_id)
Loop
retprice = GetPrice(strJson, index_write)
price = retprice
If price < highprice Then
highprice = price
End If
GET_JITA_BUY = highprice
End Function

Function GetPrice(inputStr As String, start_find As Long) As String
Dim Find1 As Long
Dim Find2 As Long
Dim GetString As String
Find1 = InStr(start_find, inputStr, "price")
Find2 = InStr(Find1, inputStr, "range")
GetString = Mid(inputStr, Find1 - 1, Find2 - Find1 - 1)
GetPrice = Mid(GetString, 9)
End Function
