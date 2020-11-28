Set WshShell  = CreateObject("WScript.Shell")

Const SrchLineCN = "Common Name,Real Address,Bytes Received,Bytes Sent,Connected Since"
Const SrchLineRT = "ROUTING TABLE"
Const SrchLineGS = "GLOBAL STATS"

Const SrchLineCL = "CLIENT_LIST"
Const SrchLineR_T = "ROUTING_TABLE"

const Name = 0
const RealAdds = 1
const BytesRx = 2
const BytesTx = 3
const ConnectedTime = 4
const VirtualAdds = 5
const VirtualAdds6 = 6

const Name2 = 1
const RealAdds2 = 2
const BytesRx2 = 5
const BytesTx2 = 6
const ConnectedTime2 = 7
const VirtualAdds2 = 3
const VirtualAdds62 = 4

Dim FileName

Dim arrFileAll

Dim CommonName
Dim RealAddress
Dim BytesReceived
Dim BytesSent
Dim ConnectedSince
Dim VirtualAddress
Dim VirtualAddress6
Dim TimeBefore : TimeBefore = 0
Dim BytesReceivedBefore
Dim BytesSentBefore
Dim SpeedReceived
Dim SpeedSent
Dim TimeNow : TimeNow = 0
Dim MaxBandwidth : MaxBandwidth = 0
Dim UnitSize
Dim LogFormat : LogFormat = 1


Dim NmbrLineCN : NmbrLineCN = 1
Dim NmbrLineRT : NmbrLineRT = 1
Dim NmbrLineGS : NmbrLineGS = 1

Sub forceCScriptExecution
    Dim Arg, Str
    If Not LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
            If InStr( Arg, " " ) Then Arg = """" & Arg & """"
            Str = Str & " " & Arg
        Next
        CreateObject( "WScript.Shell" ).Run "cscript //nologo """ & WScript.ScriptFullName & """ " & Str
        WScript.Quit
    End If
End Sub

Sub DefaultFileName()
	Set WshShell  = CreateObject("WScript.Shell")
	ProgramFile = WshShell.ExpandEnvironmentStrings("%ProgramFiles%")
	FileName = ProgramFile & "\OpenVPN\log\openvpn-status.log"
	Set WshShell  = Nothing
End Sub

Sub CLS()
    Const WshRunning  = 0
    Const WshFinished = 1
    Const WshFailed   = 2
    
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set ObjExec = WshShell.Exec("mode.com con: lines=0")
    With ObjExec
        If ObjExec.Status <> WshFailed Then
            If ObjExec.Status = WshRunning Then
                Do Until ObjExec.Status = WshFinished
                    ObjExec.StdOut.ReadAll
                    ObjExec.StdErr.ReadAll
                    WScript.Sleep 100
                Loop
            End If
        End If
    End With
	Set WshShell = Nothing
	Set ObjExec = Nothing
End Sub

Sub GetArguments() 
	Set objArgs = Wscript.Arguments
	
	Select case objArgs.Count
		case 0
			WScript.Echo "The default parameters are used:"
			WScript.Echo "To override the parameters, run the script with the following options"
			WScript.Echo "SYNOPSIS"
			WScript.Echo "ovpnstatus.vbs [status-file] [status-version]"
			Wscript.Echo "Openvpn-status file is: " & FileName
			Wscript.Echo "status-version is: 1"
			WScript.Echo
		case 1
			FileName = objArgs(0)
			Wscript.Echo "Openvpn-status file is: " & FileName
		case 2
			FileName = objArgs(0)
			LogFormat = objArgs(1)
			Wscript.Echo "Openvpn-status file is: " & FileName
	End select
End Sub

Function PrecessFile (aFileName)
	Const ForReading = 1
	Dim arrStatusFile()
	Dim FileLine
	dim i : i = 0
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set StatusFile = objFSO.OpenTextFile(aFileName, ForReading)
	StatusFile.SkipLine
	
	Do While StatusFile.AtEndOfStream <> true
		FileLine = StatusFile.Readline
		ReDim Preserve arrStatusFile(i)
		arrStatusFile(i) = FileLine
		i = i + 1
	Loop
	PrecessFile = arrStatusFile
	StatusFile.Close
	Set StatusFile = Nothing
	Set objFSO = Nothing
End Function

Function SearchLineNumber(aSrchLine, aArray)
	For i = 0 to UBound(aArray)
		If InStr(aArray(i), aSrchLine) > 0 Then
			SearchLineNumber = i + 1
		End if
	Next
End Function

Function GetRowInfo(aRowNbr, aArray)
	dim arrLine
	dim arrRow()
	dim arrRealIP
	dim k,j,n
	
	Do Until IsArrayDimmed(arrRow) <> 0
		Wscript.Sleep 500
		If (aRowNbr = 0) or (aRowNbr = 2) or (aRowNbr = 3) or (aRowNbr = 4) then
			k = 0
			For i = NmbrLineCN to NmbrLineRT - 2
				arrLine = Split(aArray(i) , ",")
				ReDim Preserve arrRow(k)
				arrRow(k) = arrLine(aRowNbr)
				k = k + 1
			Next
		ElseIf (aRowNbr = 1) then
			n = 0
			For i = NmbrLineCN to NmbrLineRT - 2
				arrLine = Split(aArray(i), ",")
				ReDim Preserve arrRow(n)
				arrRealIP = Split(arrLine(aRowNbr), ":")
				arrRow(n) = arrRealIP(0)
				n = n + 1
			Next
		ElseIf (aRowNbr = 5) then
			j = 0 
			For i = NmbrLineRT + 1 to NmbrLineGS - 2
				arrLine = Split(aArray(i) , ",")
				ReDim Preserve arrRow(j)
				arrRow(j) = arrLine(0)
				j = j + 1
			Next
		End If
		If (aRowNbr = 6) then
			k = 0
			For i = NmbrLineCN to NmbrLineRT - 2
				ReDim Preserve arrRow(k)
				arrRow(k) = "N/A"
				k = k + 1
			Next
		End If
	Loop
	GetRowInfo = arrRow
End Function

Function GetRowInfo2 (aRowNbr, aArray)
	dim arrLine
	dim arrRow()
	dim k
	dim aDelim
	
	If LogFormat = 2 then
		aDelim = Chr(44)
	ElseIf LogFormat = 3 then
		aDelim = vbTab
	End If
	
	Do Until IsArrayDimmed(arrRow)
		Wscript.Sleep 500
		k = 0
		For i = 0 to UBound(aArray)
			arrLine = Split(aArray(i), aDelim)
			If arrLine(0) = SrchLineCL then
				ReDim Preserve arrRow(k)
				If aRowNbr = 4 AND len(arrLine(aRowNbr)) = 0 then
					arrRow(k) = "N/A"
				Else
					arrRow(k) = arrLine(aRowNbr)
				End If
				k = k + 1 
			End If
		Next
	Loop
	GetRowInfo2 = arrRow
End Function

Function BytToMib(aBytes)
	dim m
	dim g
	
	if aBytes =< 1073741824 then
		m = aBytes / 1024 ^ 2
		UnitSize = "MiB"
		BytToMib = Round(m,1)
	Else
		g = aBytes / 1024 ^ 3
		BytToMib = Round(g,1)
		UnitSize = "GiB"
	End If
End Function

Function BpsToMbps (aBps)
	dim Mbps
	
	Mbps = aBps * 0.000008
	BpsToMbps = Round(Mbps,3)
End Function

Sub PrintSpeedPercentBar (aSpeed)
	dim SpeedPercent
	dim bandwidth
	dim chrTilde : chrTilde = 24
	dim chrNumber
	
	If aSpeed > MaxBandwidth then
		MaxBandwidth = aSpeed
		bandwidth = MaxBandwidth + (MaxBandwidth * 5 / 100)
	else
		bandwidth = MaxBandwidth + (MaxBandwidth * 5 / 100)
	End If
	
	If (aSpeed > 0) and (MaxBandwidth > 0) then
		SpeedPercent = aSpeed * 100 / bandwidth
		chrNumber = Round(chrTilde * SpeedPercent / 100, 1)
		WScript.StdOut.Write String(chrNumber, Chr(35)) & String(chrTilde - chrNumber, Chr(126)) & " " & aSpeed & " Mb/s"
	Else
		WScript.StdOut.Write aSpeed & " Mb/s"
	End If
End Sub

Function ClcSpeed (aByteAfter, aByteBefore, aTimeAfter, aTimeBefore)
	dim DeltaBytes
	dim DeltaTime
	dim arrSpeed()
	If IsArrayDimmed(aByteBefore) <> 0 then
		For i = 0 to UBound(aByteAfter)
			DeltaBytes = aByteAfter(i) - aByteBefore(i)
			DeltaTime = aTimeAfter - aTimeBefore
			ReDim Preserve arrSpeed(i)
			If DeltaTime > 0 then 
				arrSpeed(i) = Round (DeltaBytes / DeltaTime, 2)
			else
				arrSpeed(i) = 0
			End If
		Next
	Else
		For i = 0 to UBound(aByteAfter)
			ReDim Preserve arrSpeed(i)
			arrSpeed(i) = 0
		Next
	End If
	ClcSpeed = arrSpeed
End Function

Function IsArrayDimmed(arr)
	dim ub 
	IsArrayDimmed = False
	If IsArray(arr) Then
		On Error Resume Next
		ub = UBound(arr)
		If (Err.Number = 0) And (ub >= 0) Then 
			IsArrayDimmed = True
		End If
	End If  
End Function

Sub PrintResult
	dim columnName1 : columnName1 = "Common Name"
	dim columnName2 : columnName2 = "Real IP"
	dim columnName3 : columnName3 = "Virtual IPv4"
	dim columnName4 : columnName4 = "Virtual IPv6"
	dim columnName5 : columnName5 = "Bytes From Client:"
	dim columnName6 : columnName6 = "Bytes To Client:  "
	dim intSpace : intSpace = 25
	
	WScript.Echo "== "& Now & " " & String(50, "=")
	
	
	WScript.Echo columnName1 & Space(intSpace - Len(columnName1) - 4) & _
				columnName2 & Space(intSpace - Len(columnName2) - 2) & _
				columnName3 & Space(intSpace - Len(columnName3) - 8) & _
				columnName4 & Space(intSpace - Len(columnName4) - 2)
	For i = 0 to UBound(CommonName)
		WScript.Echo Left(CommonName(i),20) & Space(intSpace - Len(CommonName(i)) - 4) & _
				RealAddress(i) & Space(intSpace - Len(RealAddress(i)) - 2) & _
				VirtualAddress(i) & Space(intSpace - Len(VirtualAddress(i)) - 8) & _
				VirtualAddress6(i) & Space(intSpace - Len(VirtualAddress6(i)) - 8)
		WScript.StdOut.Write "Receiv Speed: " 
		Call PrintSpeedPercentBar (BpsToMbps(SpeedReceived(i)))
		Wscript.Echo " | " & columnName5 & " " & BytesReceived(i) & "  " & BytToMib(BytesReceived(i)) & "" & UnitSize & Space(intSpace - Len(BytesReceived(i)) - 11)
		WScript.StdOut.Write "Send Speed:   " 
		Call PrintSpeedPercentBar (BpsToMbps(SpeedSent(i)))
		Wscript.Echo " | " & columnName6 & " " & BytesSent(i) & "  " & BytToMib(BytesSent(i)) & "" & UnitSize & Space(intSpace - Len(BytesSent(i)))
		Wscript.Echo "Connected Since: " & ConnectedSince(i)
		Wscript.Echo
	Next
End Sub

Call forceCScriptExecution()

Call DefaultFileName()

Do while true
	Call GetArguments

	arrFileAll = PrecessFile(FileName)

	TimeBefore = TimeNow
	BytesReceivedBefore = BytesReceived
	BytesSentBefore = BytesSent

	If LogFormat = 1 then 
		NmbrLineCN = SearchLineNumber(SrchLineCN, arrFileAll)
		NmbrLineRT = SearchLineNumber(SrchLineRT, arrFileAll)
		NmbrLineGS = SearchLineNumber(SrchLineGS, arrFileAll)
	
		CommonName = GetRowInfo(Name, arrFileAll)
		RealAddress = GetRowInfo(RealAdds, arrFileAll)
		BytesReceived = GetRowInfo(BytesRx, arrFileAll)
		BytesSent = GetRowInfo(BytesTx, arrFileAll)
		ConnectedSince = GetRowInfo(ConnectedTime, arrFileAll)
		VirtualAddress = GetRowInfo(VirtualAdds, arrFileAll)
		VirtualAddress6 = GetRowInfo(VirtualAdds6, arrFileAll)
	ElseIf LogFormat = 2 or LogFormat = 3 then
		CommonName = GetRowInfo2(Name2, arrFileAll)
		RealAddress = GetRowInfo2(RealAdds2, arrFileAll)
		BytesReceived = GetRowInfo2(BytesRx2, arrFileAll)
		BytesSent = GetRowInfo2(BytesTx2, arrFileAll)
		ConnectedSince = GetRowInfo2(ConnectedTime2, arrFileAll)
		VirtualAddress = GetRowInfo2(VirtualAdds2, arrFileAll)
		VirtualAddress6 = GetRowInfo2(VirtualAdds62, arrFileAll)
	End If
	
	TimeNow = Timer
	SpeedReceived = ClcSpeed (BytesReceived, BytesReceivedBefore, TimeNow, TimeBefore)
	SpeedSent = ClcSpeed (BytesSent, BytesSentBefore, TimeNow, TimeBefore)
	
	call PrintResult()
	Wscript.Sleep 60000
	call CLS
Loop

Set WshShell = Nothing