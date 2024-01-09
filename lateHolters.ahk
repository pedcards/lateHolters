#Requires AutoHotkey v2

SetWorkingDir(A_ScriptDir)

inFile := "unsigned.xlsx"
checks := "Sea Bellevue North|Central WA"

xl := ComObjGet(A_WorkingDir "\" inFile)

for sheet in xl.worksheets
{
	if !(sheet.Name~=checks) {																	; Skip worksheets if not in "checks"
		continue
	}
	readSheet(sheet)
}

readSheet(sheet) {
	colArr := ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"]
	colTitle := ["Patient Name","MRN","Date","TRRIQ done","EP read"]
	colIdx :=	["Name","MRN","Date","TRRIQ","Read"]
	colName := colMRN := colDate := colTRRIQ := colRead := ""

	orderID := true

	while (orderID) {																			; read rows
		rowNum := A_Index
		rowArr := Map()

		loop 																					; read columns
		{
			colNum := A_Index
			cel := sheet.Range(colArr[colNum] rowNum).Value
			if (cel="ORD ID") {
				titleRow := rowNum
			}
			if (rowNum=titleRow) {
				if (x := ObjHasValue(colTitle,cel)) {
					col%colIdx[x]% := colNum
				}
				if (cel="") {
					maxcol := colNum
					break
				}
				continue
			}
			if (colnum=maxcol) {
				break
			}
			if (colNum=colName) {
				rowArr.Name := cel
			}
			if (colNum=colMRN) {
				rowArr.MRN := RegExReplace(cel,"E")
			}
			if (colNum=colDate) {
				rowArr.Date := cel
			}
		}
		if (rowNum=titleRow) {
			continue
		}
		x := checkArchive(rowArr)
	
	}
		
	
}

checkArchive(rowArr) {

}

ObjHasValue(aObj, aValue, rx:="") {
; modified from http://www.autohotkey.com/board/topic/84006-ahk-l-containshasvalue-method/	
	for key, val in aObj
		if (rx) {
			if (aValue="") {															; null aValue in "RX" is error
				return false
			}
			if (val ~= aValue) {														; val=text, aValue=RX
				return key
			}
			if (aValue ~= val) {														; aValue=text, val=RX
				return key
			}
		} else {
			if (val = aValue) {
				return key
			}
		}
	return false
}
