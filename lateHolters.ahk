#Requires AutoHotkey v2

SetWorkingDir(A_ScriptDir)

yXml := ComObject("Msxml2.DOMDocument.6.0")
yXml.load("worklist.xml")
yArch := ComObject("Msxml2.DOMDocument.6.0")
yArch.load("archive.xml")

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
	colTitle := ["Patient Name","MRN","Date","TRRIQ done","EP read","Note"]
	colIdx :=	["Name","MRN","Date","TRRIQ","Read","Note"]
	colName := colMRN := colDate := colTRRIQ := colRead := colNote := ""

	loop 																						; read rows
	{
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
			if (colNum=colDate) {
				rowArr.Date := cel
			}
		}
		if (rowNum=titleRow) {
			continue
		}
		if (rowArr.Name="") {
			break
		}
		if (processed := checkArchive(rowArr)) {
			sheet.Range(colArr[colTRRIQ] rowNum).Value := processed.TRRIQ
			sheet.Range(colArr[colRead] rowNum).Value := processed.EP
			sheet.Range(colArr[colNote] rowNum).Value := processed.Note
		}
		if (rowNum=300) {
			break
		}
	}
		
	
}

checkArchive(rowArr) {
	global yXml, yArch

	archivePath := ".\ArchiveHL7"
	tempfilesPath := "\\childrens\files\HCHolterDatabase\PROD\TRRIQ\tempfiles"

	name := stRegX(rowArr.Name,"",1,0,",",1)
	date := ParseDate(rowArr.Date).YMD
	pattern := "TRRIQ_ORU_" name "_*" ;date "*"
	archList := tempList := ""
	trDate := epDate := note := ""
	res := Map()

	/*	Find archived HL7 with name and order date
	*/
	loop files, archivePath "\" pattern
	{
		RegExMatch(A_LoopFileName,"_(\d+)_@",&dt)
		if (DateDiff(dt.1,date,"Days") < 0) {											; skip files with date before order date
			continue
		}

		archList .= A_LoopFileName "`n"
	}
	archList := Sort(archList)
	loop parse Trim(archList,"`n`r"), "`n"
	{
		filepath := archivePath "\" A_LoopField
		RegExMatch(A_LoopField,"_(\d+)_@",&dt)
		txt := FileRead(filepath)
		PID := StrSplit(stregX(txt,"PID",0,0,"\R+",0),"|")
		PID.Name := StrReplace(PID[6],"^",",")
		fuzz := FuzzySearch(rowArr.Name,PID.Name)
		trDate := ParseDate(stRegX(txt,"TRRIQ\|HS\|\|",1,1,"\|\|ORU",1)).MDY			; get date MA processed
		epDate := ParseDate(FileGetTime(filepath)).MDY									; get date ORU created (EP signed)
		if (DateDiff(dt.1,date,"Days")>60) {											; skip files if study too recent (prob new order)
			continue
		}
		if (date != dt.1) {
			note := "Order date " rowArr.Date ", Study date " ParseDate(dt.1).MDY
		}
		if (fuzz<0.1) {
			return {TRRIQ:trDate,EP:epDate,Note:note}
		}
		if (fuzz<0.2) {
			ask := MsgBox("cel name=" rowArr.Name "`nORU name=" PID.Name
						,"Name match?","0x23")
			if (ask="Yes") {
				note .= "cel name=" rowArr.Name "`nORU name=" PID.Name ". " note
				return {TRRIQ:trDate,EP:epDate,Note:note}
			}
		}
	}

	/*	No result found in ArchiveHL7, look for evidence of processing
	*/
	loop files, tempfilesPath "\*" name " *.csv", "R"
	{
		filepath := A_LoopFileFullPath
		RegExMatch(A_LoopFileName,"i)(\d+) (.*?) (\d{1,2}-\d{1,2}-\d{2,4}).csv",&fnam)
		ddiff := DateDiff(ParseDate(fnam.3).YMD,date,"Days")
		if (ddiff < 0) {															; file is before order date
			continue
		}
		if (ddiff > 60) {															; skip files if study too recent (prob new order)
			continue
		}
		tempList .= A_LoopFileFullPath "`n"
	}
	tempList := Trim(Sort(tempList),"`n`r")
	loop parse tempList, "`n"
	{
		filepath := A_LoopField
		res[A_LoopField] := Map()
		res["col"] := Map()
		loop read filepath
		{
			linenum := A_Index
			loop parse A_LoopReadLine, "CSV"
			{
				colnum := A_Index
				if (linenum=1) {
					res["col"][colnum] := A_LoopField
					res[A_LoopField] := ""
				} else {
					colidx := res["col"][colnum]
					res[colidx] := A_LoopField
				}
			}
		}
		if (rowArr.Date != res["dem-Test_date"]) {
			note := "Order date " rowArr.Date ", Study date " res["dem-Test_date"]
		}
		dName := res["dem-Name_L"] "," res["dem-Name_F"]
		fuzz := FuzzySearch(rowArr.Name, dName)
		if (fuzz<0.1) {
			return {TRRIQ:res["dem-Test_date"],EP:epDate,Note:note}
		}
		if (fuzz<0.2) {
			ask := MsgBox("cel name=" rowArr.Name "`nORU name=" dName
						,"Name match?","0x23")
			if (ask="Yes") {
				note .= "cel name=" rowArr.Name "`nORU name=" dName ". " note
				return {TRRIQ:res["dem-Test_date"],EP:epDate,Note:note}
			}
		}
	}
}

ParseDate(x) {
	mo := ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
	moStr := "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
	dSep := "[ \-\._/]"
	date := {yyyy:"",mmm:"",mm:"",dd:"",date:""}
	time := {hr:"",min:"",sec:"",days:"",ampm:"",time:""}

	x := RegExReplace(x,"[,\(\)]")

	if (x~="\d{4}.\d{2}.\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z") {
		x := RegExReplace(x,"[TZ]","|")
	}
	if (x~="\d{4}.\d{2}.\d{2}T\d{2,}") {
		x := RegExReplace(x,"T","|")
	}

	if RegExMatch(x,"i)(\d{1,2})" dSep "(" moStr ")" dSep "(\d{4}|\d{2})",&d) {			; 03-Jan-2015
		date.dd := zdigit(d[1])
		date.mmm := d[2]
		date.mm := zdigit(objhasvalue(mo,d[2]))
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(\d{4})[\-\.](\d{2})[\-\.](\d{2})\b",&d) {					; 2015-01-03
		date.yyyy := d[1]
		date.mm := zdigit(d[2])
		date.mmm := mo[d[2]]
		date.dd := zdigit(d[3])
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr "|\d{1,2})" dSep "(\d{1,2})" dSep "(\d{4}|\d{2})",&d) {	; Jan-03-2015, 01-03-2015
		date.dd := zdigit(d[2])
		date.mmm := objhasvalue(mo,d[1]) 
			? d[1]
			: mo[d[1]]
		date.mm := objhasvalue(mo,d[1])
			? zdigit(objhasvalue(mo,d[1]))
			: zdigit(d[1])
		date.yyyy := (d[3]~="\d{4}")
			? d[3]
			: (d[3]>50)
				? "19" d[3]
				: "20" d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"i)(" moStr ")\s+(\d{1,2}),?\s+(\d{4})",&d) {					; Dec 21, 2018
		date.mmm := d[1]
		date.mm := zdigit(objhasvalue(mo,d[1]))
		date.dd := zdigit(d[2])
		date.yyyy := d[3]
		date.date := trim(d[0])
	}
	else if RegExMatch(x,"\b(19\d{2}|20\d{2})(\d{2})(\d{2})((\d{2})(\d{2})(\d{2})?)?\b",&d)  {	; 20150103174307 or 20150103
		date.yyyy := d[1]
		date.mm := d[2]
		date.mmm := mo[d[2]]
		date.dd := d[3]
		if (d[1]) {
			date.date := d[1] "-" d[2] "-" d[3]
		}
		
		time.hr := d[5]
		time.min := d[6]
		time.sec := d[7]
		if (d[5]) {
			time.time := d[5] ":" d[6] . strQ(d[7],":###")
		}
	}

	if RegExMatch(x,"i)(\d+):(\d{2})(:\d{2})?(:\d{2})?(.*)?(AM|PM)?",&t) {				; 17:42 PM
		hasDays := (t[4]) ? true : false 											; 4 nums has days
		time.days := (hasDays) ? t[1] : ""
		time.hr := trim(t[1+hasDays])
		time.min := trim(t[2+hasDays]," :")
		time.sec := trim(t[3+hasDays]," :")
		if (time.min>59) {
			time.hr := floor(time.min/60)
			time.min := zDigit(Mod(time.min,60))
		}
		if (time.hr>23) {
			time.days := floor(time.hr/24)
			time.hr := zDigit(Mod(time.hr,24))
			DHM:=true
		}
		time.ampm := trim(t[5])
		time.time := trim(t[0])
	}

	return {yyyy:date.yyyy, mm:date.mm, mmm:date.mmm, dd:date.dd, date:date.date
			, YMD:date.yyyy date.mm date.dd
			, YMDHMS:date.yyyy date.mm date.dd zDigit(time.hr) zDigit(time.min) zDigit(time.sec)
			, MDY:date.mm "/" date.dd "/" date.yyyy
			, MMDD:date.mm "/" date.dd 
			, hrmin:zdigit(time.hr) ":" zdigit(time.min)
			, days:zdigit(time.days)
			, hr:zdigit(time.hr), min:zdigit(time.min), sec:zdigit(time.sec)
			, ampm:time.ampm, time:time.time
			, DHM:zdigit(time.days) ":" zdigit(time.hr) ":" zdigit(time.min) " (DD:HH:MM)" 
			, DT:date.mm "/" date.dd "/" date.yyyy " at " zdigit(time.hr) ":" zdigit(time.min) ":" zdigit(time.sec) }
}

zDigit(x) {
; Returns 2 digit number with leading 0
	return SubStr("00" x, -2)
}

strQ(var1,txt,null:="") {
/*	Print Query - Returns text based on presence of var
	var1	= var to query
	txt		= text to return with ### on spot to insert var1 if present
	null	= text to return if var1="", defaults to ""
*/
	return (var1="") ? null : RegExReplace(txt,"###",var1)
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

#Include strx2.ahk
#Include xml2.ahk
#Include sift3.ahk