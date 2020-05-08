Const xlUp = -4162

Function excelGetMaxRow(objSt,byVal strCol,byVal lngRow)
	dim lngRowMax
	lngRowMax = objSt.rows.count
	lngRowMax = objSt.Range(strCol & lngRowMax).End(xlUp).Row
	if lngRow > lngRowMax then
		lngRowMax = lngRow
	end if
	excelGetMaxRow = lngRowMax
End Function
