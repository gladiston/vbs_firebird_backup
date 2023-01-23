' Windows Script Host - VBScript
'-----------------------------------------------------------------
' Nome: vbs_firebird_backup.vbs
' Proposito : Realização de Backup do Servidor FB
' By: Gladiston Santana (gladiston.santana em gmail.com)
' Copyright: (c) Jun 2011, Todos os direitos reservados!
'-----------------------------------------------------------------

' Trata-se de um script vbs (Windows Script Host - VBScript) com o propósito 
' de fazer backup de base de dados FirebirdSQL 3+. Use-o conjuntamente com o agendador 
' de tarefas no servidor para automatizar o backups. também é capaz de 
' depositar os arquivos numa pasta remota, ex: \\servidor\compartilhamento.
' Modo de usar:
' vbs_firebird_backup.vbs \\server\bak "C:\dados\banco1.FDB" "C:\dados\banco2.FDB" "C:\dados\banco3.FDB"
' 
' Parametro #1: Refere-se ao destino do backup, pode ser uma pasta local ou pasta
'  remota do tipo UNC como \\servidor\compatilhamento. Se usar uma pasta remota 
'  junto com o agendador de tarefas, programe o agendador de tarefas para rodar a 
'  tarefa sob um usuario que tenha permissão sobre a pasta remota e também que 
'  as variaveis de ambiente deste usuário contenha ISC_USER e ISC_PASSWORD. 
'  Um alerta importante, se estiver num host que tenha serviços de sincronização 
'  com a nuvem como o onedrive, gdrive,... veja se o tempo de sincronização não 
'  compromete o backup, pois se acontecer um sinistro com o servidor antes que 
'  a sincronização termine e então terá um belo problema nas mãos.
' Parametro #2..N: Todos os parametros seguintes referem-se aos arquivos de dados que 
'  terão o seu backup realizado, mas atenção que devem estar em aspas duplas. 
'  Recomendo que caso opte por varios bancos de uma só vez que então coloque os bancos 
'  mais prioritarios primeiros. Caso use o agendador de tarefas do Windows com este 
'  script, crie uma programação onde os bancos mais importantes tenham um intervalo 
'  entre backups menor e os menos importantes com intervalos maiores, 
'  essa é a premissa de dividir para conquistar.
' Este script usa as variaveis de ambiente ISC_USER e ISC_PASSWORD para saber qual 
'  usuario e senha que deverá ser usada para a realização do backup. 
'  Se quiser fazer diferente, voce pode modificar as referencias:
'   fdb_server="localhost"
'   fdb_username= oWS.ExpandEnvironmentStrings("%ISC_USER%")
'   fdb_password= oWS.ExpandEnvironmentStrings("%ISC_PASSWORD%")
' Porém ao modificá-las e colocar valores explicitos voce estaria sendo imprudente, 
'   pois se este script vazar, pessoas inescrupulosas poderiam usar essas informações 
'   explicitadas para invadir o seu sistema.
'
' Este script tem suporte a voz(de bêbado) para indicar verbalmente quando inicia 
'  e quando termina, mas para ser sincero, não acho que isso seja importante 
'  especialmente em servidores, por isso caso queira desligar troque de 
' True para False na linha:
'  bWantVoice=True

' Variaveis reservadas que nao podem ser redeclaradas: 
Dim sDatabase
Dim sResultado
Dim sFB_PATH
Dim fdb_server
Dim fdb_username
Dim fdb_password
Dim gbak
Dim sInicio
Dim sRoot
Dim sDestino
Dim sOriLogFile
Dim objFolder
Dim objFile
Dim sBackupVolName
Dim sLogFile
Dim sMensagem
Dim sEscondeSenha1
Dim sEscondeSenha2
Dim bSemParar
Dim sTempFolder
Dim bWantVoice
Dim iElimina_Apos_Dias

Dim sToday_Year
Dim sToday_Month
Dim sToday_Day

Dim oFS
Dim oWS
Dim oWN
Dim oArgs 

Set oWS = WScript.CreateObject("WScript.Shell")
Set oWN = WScript.CreateObject("WScript.Network")
Set oFS = WScript.CreateObject("Scripting.FileSystemObject")
Set oArgs = Wscript.Arguments

' API de Voz do Windows (Windows 7+)
Set VOZ = CreateObject("sapi.spvoice")

' Captura ano, mes e dia. Onde mes e dia tem 2 digitos.
sToday_Year=Year(DateValue(Now()))
sToday_Month=Month(DateValue(Now()))
if Len(sToday_Month)=1 then sToday_Month= "0" & sToday_Month
sToday_Day=Day(DateValue(Now()))
if Len(sToday_Day)=1 then sToday_Day= "0" & sToday_Day
  
' Mensagem Texto para ser usada em logs e afins
sMensagem=""

' Pasta temporaria
sTempFolder = oFS.GetSpecialFolder(2)

' Arquivo de log
sLogFile=sTempFolder & "\backup-firebird-" & DataExtenso(Now(),False) & ".log"

' Copia sem paradas para dar [OK] ?
bSemParar=True

' Se não quiser voz, troque o parametro True Por False
bWantVoice=False

' Quantos dias se passará para um backup considerar expirado e poderá ser excluido
iElimina_Apos_Dias=180

' Detectando localizacao do FB3 e/ou FB4
sFB_PATH=FB_WhereIsFirebird("", True, False)

' Se nao foi encontrado entao cai fora 
If (sFB_PATH="") Then   
  Call LimpezaESair
End If 

gbak=sFB_PATH & "\gbak.exe"

If Not oFS.FileExists(gbak) Then
  Call LimpezaESair
End If

' sBackupVolName define mais uma subpasta no destino indicado
sBackupVolName=""

' Usuario e senha para conectar o firebird
fdb_server="localhost"
fdb_username= oWS.ExpandEnvironmentStrings("%ISC_USER%")
fdb_password= oWS.ExpandEnvironmentStrings("%ISC_PASSWORD%")


sEscondeSenha1=fdb_password
' Data atual no formato AAAAMMDDHHMMSS
sInicio = DataExtenso(Now(),True)

'Destino do Backup
sRoot = wscript.arguments(0)
if sBackupVolName<>"" then
  sRoot = sRoot & "\" & sBackupVolName
End If 
 
If Not oFS.FolderExists(sRoot) Then 
  Set objFolder = oFS.CreateFolder(sRoot)
  Set objFolder = Nothing  
End If

' Acrescenta a unidade de destino, o ano ...\2023
sDestino = sRoot & "\" & sToday_Year
If Not oFS.FolderExists(sDestino) Then 
  Set objFolder = oFS.CreateFolder(sDestino)
  Set objFolder = Nothing 
End If

' Acrescenta a unidade de destino, o mes ...\2023\01
sDestino = sDestino & "\" & sToday_Month
If Not oFS.FolderExists(sDestino) Then 
  Set objFolder = oFS.CreateFolder(sDestino)
  Set objFolder = Nothing 
End If

' Acrescenta a unidade de destino, o dia ...\2023\01\01
sDestino = sDestino & "\" & sToday_Day & "-" & WeekdayName(Weekday(Now()),True,1)
If Not oFS.FolderExists(sDestino) Then 
  Set objFolder = oFS.CreateFolder(sDestino)
  Set objFolder = Nothing 
End If

' Nao posso prosseguir se não foi informado um parametro contendo os bancos
if (wscript.arguments.count < 2) then
   AddToLog("Modo de Usar: vbs_firebird_backup.vbs \\server\destino c:\db\banco1.fdb c:\db\banco2.fdb c:\db\banco3.fdb")
   Call LimpezaESair
end if


' Se a pasta de destino nao foi criada então cai fora
If not oFS.FolderExists(ExtractPath(sDestino)) Then 
  AddToLog("Não foi possivel criar a pasta: " & sDestino)
  Call LimpezaESair 
End If  

' O log até essa esse ponto sera transferido para um novo local que é mais permanente
sOriLogFile=sDestino+"\" & sBackupVolName & "-" & DataExtenso(Now(),False) & ".log"
If oFS.FileExists(sLogFile) Then 
  oFS.CopyFile sLogFile, sOriLogFile, True
  oFS.DeleteFile sLogFile
End If
sLogFile= sOriLogFile 

' Definindo os bancos que terão o seu backup, basicamente cada argumento é um arquivo
'   a ser feito o backup
'--------------------
' Inicio do programa
'--------------------
' Fala que o backup esta sendo iniciado
if VoiceAPI_Installed(bWantVoice)=True Then
  VOZ.Speak "Backup do Firebird esta sendo iniciado agora"
End If  

If not bSemParar Then
	WScript.Echo "Prestes a iniciar o backup:" & vbCrlf & vbCrlf & _
		  "Firebird: " & sFB_PATH & vbCrlf & _
		  "Firebird Server: " & fdb_server & vbCrlf & _
		  "Firebird User: " &  fdb_username & vbCrlf & _
		  "Firebird Password: " &  Len(fdb_password) & " digitos" & vbCrlf & _
		  "Destino Root: " & sRoot & vbCrlf & _
		  "Destino completo: " & sDestino & vbCrlf & _
		  "Log File: " & sLogFile & vbCrlf & _
		  "Clique OK para prosseguir" 
End If

For Each sDatabase in oArgs
	If oFS.FileExists(sDatabase) And sDatabase<>sRoot Then 
       Call DoBackup(sDatabase)
	End If  
Next

' Limpeza de backups expirados
if iElimina_Apos_Dias > 0 Then
  LimparArquivosBackupsAntigos sRoot, iElimina_Apos_Dias
End If

' Fala que o backup foi concluido
if VoiceAPI_Installed(bWantVoice)=True Then
  VOZ.Speak "Backup do Firebird concluído, observe os logs em " & sDestino
End If  
'--------------------
' Finaliza o programa
'--------------------
Call LimpezaESair

'-----------------------------------------------------------------
' SubRotinas necessárias para executar apenas este script
'-----------------------------------------------------------------
Sub DoBackup(sOrigem)
  Dim sDestino_real
  Dim sDestino_temp
  Dim sArq
  Dim sCmd
  sArq=GetFileBaseName(sOrigem)
  sLogFile=sDestino & "\" & sArq & "-" & sInicio & ".log"
  sDestino_real=sDestino & "\" & sArq & "-" & sInicio & ".fbk"
  sDestino_temp=sTempFolder & "\" & sArq & "-" & sInicio & ".fbk"
  sCmd = Chr(34) & gbak & Chr(34) & " -v -b -t " &_
    " -user " & Chr(34) & fdb_username & Chr(34) &_ 
    " -password " & Chr(34) & fdb_password & Chr(34) & " " &_
    Chr(34) & fdb_server & ":" & sOrigem & Chr(34) & " "  &_
    Chr(34) & sDestino_temp & Chr(34) & " " &_
    " -Y " & Chr(34) & sLogFile & Chr(34) & " " 	

  ' Executando...
  sResultado = RunAsAgente(sCmd)
  
  If oFS.FileExists(sDestino_temp) Then 
    oFS.CopyFile sDestino_temp, sDestino_real, True
    oFS.DeleteFile sDestino_temp
  End If  
End Sub

' Simplesmente Finaliza o programa e reseta valores 
Sub LimpezaESair()
  Set oWS = Nothing
  Set oWN = Nothing
  Set oFS = Nothing
  WScript.Quit
End Sub

Function AddToLog(sMensagem)
  Dim objFile
  AddToLog=False
  Const ForAppending = 8
  set objFile = oFS.OpenTextFile(sLogFile, ForAppending, True)
  sMensagem = Replace(sMensagem,sEscondeSenha1,"*****") 
  sMensagem = Replace(sMensagem,sEscondeSenha2,"*****") 
  objFile.WriteLine(DataExtenso(Now(), true) & ";" & sMensagem)
  If (Err.Number <> 0) and ( bSemParar<>True ) Then
    WScript.Echo "Adicionando mensagem ao arquivo :" & sLogFile & vbCrlf & _
      "Código do Erro: " & Err.Number & vbCrlf & _
      "Código do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
      "Fonte: " &  Err.Source & vbCrlf & _
      "Descrição do Erro: " &  Err.Description
    Err.Clear
  else
    AddToLog=True
  End If
  objFile.Close 
End Function

' Extrai a data por extenso
Function DataExtenso(sData, sExibeHoras)
Dim Resultado, sYear, sMonth, sDay, sHour, sMin, sSec
  sYear=Year(DateValue(sData))
  sMonth=Month(DateValue(sData))
  sDay=Day(DateValue(sData))
  sHour=Hour(sData)
  sMin=Minute(sData)  
  sSec=Second(sData)  
  if Len(sMonth)=1 then sMonth= "0" & sMonth
  if Len(sDay)=1 then sDay= "0" & sDay
  if Len(sHour)=1 then sHour= "0" & sHour
  if Len(sMin)=1 then sMin= "0" & sMin
  if Len(sSec)=1 then sSec= "0" & sSec
  Resultado = sYear & "-" & sMonth & "-" & sDay 
  if sExibeHoras=True Then 
    Resultado=Resultado & "+" & sHour & "h" & sMin & "m" & sSec & "s"
  End If	
  DataExtenso = Resultado 
End Function

' Extrai o path de um arquivo que esteja numa string
Function ExtractPath(sFileName)
Dim strPath, strFileName, lngIndex 
  ExtractPath=sFileName
  strPath = Split(sFileName, "\")
  lngIndex = UBound(strPath)
  strFileName = strPath(lngIndex)
  strPath(lngIndex) = "" 
  ExtractPath=Join(strPath, "\")
End Function

' Extrai o nome de um arquivo que esteja numa string
Function GetFileBaseName(sFileName)
Dim strFile, sArq 
  GetFileBaseName=sFileName
  strFile = Split(sFileName, "\")
  sArq = strfile(UBound(strFile))
  strFile = Split(sArq, ".")
  sArq = strfile(LBound(strFile))
  GetFileBaseName=sArq
End Function

' Extrai a extensão de um arquivo que esteja numa string
Function GetFileExt(sFileName)
Dim strFile, sArq, sExt
  strFile = Split(sFileName, "\")
  sArq = strfile(UBound(strFile))
  strFile = Split(sArq, ".")
  sExt = strfile(UBound(strFile))
  'WScript.Echo sExt
  GetFileExt="." & sExt
End Function

Function RunAsAgente(ACMD)
  RunAsAgente = False
  If Not oFS.FileExists(ACMD) Then
	  oWS.run ACMD, 1, True  
	  'oWS.run(ACMD)
	  If Err.Number <> 0 Then
		 sMensagem = vbTab & "  Falhou :" & ACMD & vbCrlf & _
		   "Código do Erro: " & Err.Number & vbCrlf & _
		   "Código do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
		   "Fonte: " &  Err.Source & vbCrlf & _
		   "Descrição do Erro: " &  Err.Description
		 AddToLog(sMensagem)
		 if  ( bSemParar<>False ) Then WScript.Echo( sMensagem )
	  else
		 RunAsAgente=False
		 sMensagem = vbTab & "Sucesso : " & ACMD
		 AddToLog(sMensagem)	   
	  End If    
	  Err.Clear
  End If
End Function

' Elimina arquivos de uma pasta que sejam mais antigos que a 
' data atual - iDaysOld 
Sub LimparArquivosBackupsAntigos(sDirectoryPath, iDaysOld)
Dim oFolder
Dim oFileCollection
Dim oFile
Dim bPodeApagar
Dim sFileName  
Dim sDir   
  Set oFS = CreateObject("Scripting.FileSystemObject") 
  Set oFolder = oFS.GetFolder(sDirectoryPath) 
  Set oFileCollection = oFolder.Files 

  For each oFile in oFileCollection
    'No exemplo abaixo, apenas arquivos com a extensão .fbk, .log e .dat seriam apagados
    bPodeApagar=False 
    sFileName=Cstr(sDirectoryPath) & "\" & Cstr(oFile.Name)
    If LCase(Right(Cstr(oFile.Name), 3)) = "fbk" Then bPodeApagar=True
    If LCase(Right(Cstr(oFile.Name), 3)) = "dat" Then bPodeApagar=True
    If LCase(Right(Cstr(oFile.Name), 3)) = "log" Then bPodeApagar=True
    If bPodeApagar = True Then
      If oFile.DateLastModified < (Date() - iDaysOld) Then 
        oFile.Delete(True)   
        If Err.Number <> 0 Then
          sMensagem = vbTab & "  Falhou ao eliminar arquivo :" & sFileName & vbCrlf & _
            "Código do Erro: " & Err.Number & vbCrlf & _
            "Código do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
            "Fonte: " &  Err.Source & vbCrlf & _
            "Descrição do Erro: " &  Err.Description
          AddToLog(sMensagem)
        Else
          AddToLog( "Arquivo eliminado: " & sFileName)   
        End If 
      End If
    End If   
  Next 
  Set oFolder = Nothing 
  Set oFileCollection = Nothing 
  Set oFile = Nothing   
End Sub

'VoiceAPI_Installed: Detecta se o Windows atual teria suporte a voz
'Essa função é dummy ainda, pois não achei um metodo seguro para fazer
'isso que funcione desde o Windows 2003
Function VoiceAPI_Installed(bReqVoice)
  VoiceAPI_Installed=bReqVoice
  'Todo: Criar um codigo que saiba que apenas Win7+ tem suporte a voz
  'VoiceAPI_Installed=True/False
End Function

Function FB_WhereIsFirebird(sIfNotFoundReturnAs, bCheckFB3, bCheckFB4)
    Dim S
    S=""
	if (bCheckFB3=True) Then
		If ((S="") and _
		   (oFS.FolderExists("C:\Arquivos de programas\Firebird\Firebird_3_0"))) Then 
		  S="C:\Arquivos de programas\Firebird\Firebird_3_0"
		End If
		 
		If (S="") and _
			(oFS.FolderExists("C:\Program Files\Firebird\Firebird_3_0")) Then 
		   S="C:\Program Files\Firebird\Firebird_3_0"
		End If  
		 
		If (S="") and _
			(oFS.FolderExists("C:\Arquivos de programas (x86)\Firebird\Firebird_3_0")) Then 
		  S="C:\Arquivos de programas (x86)\Firebird\Firebird_3_0"
		End If
		 
		If (S="") and _
		   (oFS.FolderExists("C:\Program Files (x86)\Firebird\Firebird_3_0")) Then 
		  S="C:\Program Files (x86)\Firebird\Firebird_3_0"
		End If  
	End If	
	if (bCheckFB4=True) Then
		If ((S="") and _
		   (oFS.FolderExists("C:\Arquivos de programas\Firebird\Firebird_4_0"))) Then 
		  S="C:\Arquivos de programas\Firebird\Firebird_4_0"
		End If
		 
		If (S="") and _
			(oFS.FolderExists("C:\Program Files\Firebird\Firebird_4_0")) Then 
		   S="C:\Program Files\Firebird\Firebird_4_0"
		End If  
		 
		If (S="") and _
			(oFS.FolderExists("C:\Arquivos de programas (x86)\Firebird\Firebird_4_0")) Then 
		  S="C:\Arquivos de programas (x86)\Firebird\Firebird_4_0"
		End If
		 
		If (S="") and _
		   (oFS.FolderExists("C:\Program Files (x86)\Firebird\Firebird_4_0")) Then 
		  S="C:\Program Files (x86)\Firebird\Firebird_4_0"
		End If  
	End If	

    if S="" Then
		FB_WhereISFirebird=sIfNotFoundReturnAs
	Else
		FB_WhereISFirebird=S	
	End if

End Function
