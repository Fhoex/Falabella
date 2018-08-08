Option Explicit
Dim FSO, archivo, matriz, count, objWSh
Dim PrinterName, DriverName, PortName, HostAddress, Protocol, Queue, Server
Dim ObjFolder, ObjSubFolders, Locale(1), Path, i, BullZip, FileCSV, BreakCapture

'-------------------------------------------------------------------------------------------------
FileCSV = "PrintersDetails.csv"
BullZip = "Setup_BullzipPDFPrinter_11_7_0_2716_PRO_EXP.exe"
Path = "C:\Windows\system32\printing_admin_scripts\"
Server = "127.0.0.1"
'-------------------------------------------------------------------------------------------------

Set objWSh = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set archivo = FSO.OpenTextFile(FileCSV)
Set ObjFolder = FSO.GetFolder(Path)
Set ObjSubFolders = ObjFolder.SubFolders
count = 0
i = 0

For Each ObjFolder In ObjSubFolders
	Locale(i) = ObjFolder.Name
	i = i + 1
	Next

Do While NOT archivo.AtEndOfStream
	matriz	= split(archivo.ReadLine,",")
	PortName = "ricoh_" & matriz(0)
	HostAddress = matriz(0)
	Protocol = matriz(1)
	Queue = matriz(2)
	PrinterName = matriz(3)
	DriverName = matriz(4)
	BreakCapture = ""
	objWSh.Run BullZip & " /PRINTERNAME=""" & matriz(4) & "", 0, True
	
	Do While BreakCapture <> "X"
		BreakCapture = InputBox("Ingrese el tipo de cola (T1,T3,T4) o X para terminar",PrinterName & " SubQueue")
		If BreakCapture <> "X" Then
			If HostAddress <> "" And Protocol <> "" And Queue <> "" Then
				objWSh.Run "cscript " & Path & Locale(0) & "\prnport.vbs -l -s " & Server & " -t -r " & PortName & " -h " & HostAddress & " -o " & Protocol & " -n " & Queue,0,True
				End If
			If PrinterName <> "" And DriverName <> "" Then
				objWSh.Run "cscript " & Path & Locale(0) & "\prnmngr.vbs -a -s " & Server & " -p " & PrinterName & "_" & BreakCapture & " -m """ & DriverName & """ -r " & PortName,0,True '[-u UserName -w Password]
				End If
			If PrinterName <> "" Then
				objWSh.Run "cscript " & Path & Locale(0) & "\prncnfg.vbs -t -s " & Server & " -p " & PrinterName & "_" & BreakCapture & "" & "" & "",0,True
				End If
			count = count + 1
			For i=0 To 4 Step 1
				matriz(i) = ""
				Next
			End If
		Loop
	Loop
archivo.Close
Set archivo = Nothing
Set objWSh = Nothing
Set FSO = Nothing
WScript.Echo "Proceso Finalizado (" & count & ")"
WScript.Quit