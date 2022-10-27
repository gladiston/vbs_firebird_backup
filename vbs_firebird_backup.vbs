'' Windows Script Host - VBScript
'-----------------------------------------------------------------
' Nome: vbs_firebird_backup.vbs
' Proposito : Realiza��o de Backup do Servidor FB
' By: Gladiston Santana (gladiston.santana em gmail.com)
' Copyright: (c) Jun 2011, Todos os direitos reservados!
'-----------------------------------------------------------------

' Trata-se de um script vbs (Windows Script Host - VBScript) com o prop�sito 
' de fazer backup de base de dados FirebirdSQL. Use-o conjuntamente com o agendador 
' de tarefas no servidor. Ele faz o backup e deposita os arquivos numa pasta local ou remota.
' Modo de usar:
' vbs_firebird_backup.vbs \\server\bak "C:\dados\banco1.FDB" "C:\dados\banco2.FDB" "C:\dados\banco3.FDB"
' 
' Parametro #1: Refere-se ao destino do backup, pode ser uma pasta local 
'  ou pasta remota do tipo UNC como \servidor\comparilhamento. Se usar 
'  uma pasta remota junto com o agendador de tarefas, programe o 
'  agendador de tarefas para rodar a tarefa sob um usuario que tenha 
'  permiss�o a pasta remota. Um alerta importante, se estiver num host 
'  que tenha servi�os de sincroniza��o com a nuvem como o onedrive, 
'  gdrive,... estes servi�os n�o sabem quando o backup terminou e por 
'  essa raz�o, quando o backup iniciar-se a cada byte de backup o 
'  programa ir� querer sincronizar quando deveria esperar terminar 
'  primeira, isso n�o compromete o backup, mas tornar� a sincroniza��o 
'  para a nuvem bastante demorada.
' Parametro #2...N: Todos os parametros seguintes referem-se aos arquivos de dados 
'   que ter�o o seu backup realizado, mas aten��o que devem estasr em aspas duplas. 
'   Recomendo que caso opte por varios bancos de uma s� vez que ent�o coloque os 
'   bancos mais prioritarios primeiros. Caso use o agendador de tarefas do Windows, 
'   crie uma programa��o onde os bancos mais importantes tenham um intervalo entre 
'   backups menor e os menos importantes com intervalos maiores, essa � a premissa 
'   de dividir para conquistar.
' Este script usa a variavel de ambiente ISC_USER e ISC_PASSWORD para saber qual 
' usuario e senha dever� ser usada para a realiza��o do backup. 
' Voce pode trocar as referencias:
'   fdb_server="localhost"
'   fdb_username= oWS.ExpandEnvironmentStrings("%ISC_USER%")
'   fdb_password= oWS.ExpandEnvironmentStrings("%ISC_PASSWORD%")
' Ao modificar e colocar valores explicitos nessas variaveis voce estaria engessando 
'  o host, usuario e senha e n�o dependeria mais variaveis de ambiente, isso facilitaria, 
'  mas isto seria imprudente, pois se este script vazar, pessoas inescrupulosas 
'  poderiam usar essa informa��o para invadir o seu sistema.

' Variaveis reservadas que nao podem ser redeclaradas: 
Dim sDatabase, sResultado, sFB_PATH
Dim fdb_server, fdb_username, fdb_password, lista_fdb, gbak
Dim sInicio, sRoot, sDestino, sOriLogFile
Dim objFolder, objFile, sBackupVolName
Dim oFS, oWS, oWN
Dim q,cr,sLogFile,sSemParar,sMensagem, sEscondeSenha1, sEscondeSenha2
Dim sTempFolder
Dim oArgs 

Set oWS = WScript.CreateObject("WScript.Shell")
Set oWN = WScript.CreateObject("WScript.Network")
Set oFS = WScript.CreateObject("Scripting.FileSystemObject")
Set oArgs = Wscript.Arguments

' API de Voz do Windows (Windows 10+)
'   comente caso esteja usando Windows anteriores
Set VOICE = createobject("sapi.spvoice")

' Aspas
q=Chr(34)
' LineFeed
cr=vbCrlf
' Mensagem Texto para ser usada em logs e afins
sMensagem=""
' Pasta temporaria
sTempFolder = oFS.GetSpecialFolder(2)
' Arquivo de log
sLogFile=sTempFolder & "\backup-firebird-" & DataExtenso(Now(),False) & ".log"

' Copia sem paradas para dar [OK] ?
sSemParar=".SIM"

' Volume do Backup - Apenas discos que contenham uma pasta com o mesmo
' nome do Volume do Backup ser�o reconhecidos
sBackupVolName="bak-firebird"

' Detectando localizacao do FB3
sFB_PATH=""
If ((sFB_PATH="") and _
   (oFS.FolderExists("C:\Arquivos de programas\Firebird\Firebird_3_0"))) Then 
  sFB_PATH="C:\Arquivos de programas\Firebird\Firebird_3_0"
End If
 
If (sFB_PATH="") and _
    (oFS.FolderExists("C:\Program Files\Firebird\Firebird_3_0")) Then 
   sFB_PATH="C:\Program Files\Firebird\Firebird_3_0"
End If  
 
If (sFB_PATH="") and _
    (oFS.FolderExists("C:\Arquivos de programas (x86)\Firebird\Firebird_3_0")) Then 
  sFB_PATH="C:\Arquivos de programas (x86)\Firebird\Firebird_3_0"
End If
 
If (sFB_PATH="") and _
   (oFS.FolderExists("C:\Program Files (x86)\Firebird\Firebird_3_0")) Then 
  sFB_PATH="C:\Program Files (x86)\Firebird\Firebird_3_0"
End If  

' Se nao foi encontrado entao cai fora 
If (sFB_PATH="") Then   
  Call LimpezaESair
End If 

gbak=sFB_PATH & "\gbak.exe"

If Not oFS.FileExists(gbak) Then
  Call LimpezaESair
End If

' Volume do Backup - Apenas discos que contenham uma pasta com o mesmo
' nome do Volume do Backup ser�o reconhecidos
sBackupVolName="bak-firebird"
' Dados de usuario e senha para conectar o firebird
' para gerar nome de usuario/senha encriptado 
' use o script Test_[En/De]crypt.vbs
fdb_server="localhost"
fdb_username= oWS.ExpandEnvironmentStrings("%ISC_USER%")
fdb_password= oWS.ExpandEnvironmentStrings("%ISC_PASSWORD%")


sEscondeSenha1=fdb_password
' Data atual no formato AAAAMMDDHHMMSS
sInicio = DataExtenso(Now(),True)

'Destino do Backup
sRoot = wscript.arguments(0) & "\" & sBackupVolName
If Not oFS.FolderExists(sRoot) Then 
  Set objFolder = oFS.CreateFolder(sRoot)
  Set objFolder = Nothing  
End If

sDestino = sRoot & "\" & WeekdayName(Weekday(Now()),False,1)
If Not oFS.FolderExists(sDestino) Then 
  Set objFolder = oFS.CreateFolder(sDestino)
  Set objFolder = Nothing 
End If

' Nao posso prosseguir se n�o foi informado um parametro contendo os bancos
if (wscript.arguments.count < 2) then
   AddToLog("Modo de Usar: vbs_firebird_backup.vbs \\server\destino c:\db\banco1.fdb c:\db\banco2.fdb c:\db\banco3.fdb")
   Call LimpezaESair
end if


' Se a pasta de destino nao foi criada ent�o cai fora
If not oFS.FolderExists(ExtractPath(sDestino)) Then 
  AddToLog("N�o foi possivel criar a pasta: " & sDestino)
  Call LimpezaESair 
End If  

' O log at� essa esse ponto sera transferido para um novo local que � mais permanente
sOriLogFile=sDestino+"\" & sBackupVolName & "-" & DataExtenso(Now(),False) & ".log"
If oFS.FileExists(sLogFile) Then 
  'oFS.CopyFile sLogFile, ExtractPath(sOriLogFile), True
  oFS.CopyFile sLogFile, sOriLogFile, True
  oFS.DeleteFile sLogFile
End If
sLogFile= sOriLogFile 

' Definindo os bancos que ter�o o seu backup, basicamente cada argumento � um arquivo
'   a ser feito o backup
'--------------------
' Inicio do programa
'--------------------

For Each sDatabase in oArgs
	If oFS.FileExists(sDatabase) And sDatabase<>sRoot Then 
       Call DoBackup(sDatabase)
	End If  
Next

' Se n�o quiser voz, comente a linha abaixo.
VOICE.Speak "Backup do Firebird conclu�do, observe os logs em " & sRoot

'--------------------
' Finaliza o programa
'--------------------
Call LimpezaESair

'-----------------------------------------------------------------
' SubRotinas necess�rias para executar apenas este script
'-----------------------------------------------------------------
Sub DoBackup(sOrigem)
Dim sDestino2, sArq, sCmd
  sArq=GetFileBaseName(sOrigem)
  sLogFile=sDestino & "\" & sArq & "-" & sInicio & ".log"
  sDestino2=sDestino & "\" & sArq & "-" & sInicio & ".fbk"
  sCmd = q & gbak & q & " -v -b -t " &_
    " -user " & q & fdb_username & q &_ 
    " -password " & q & fdb_password & q & " " &_
    q & fdb_server & ":" & sOrigem & q & " "  &_
    q & sDestino2 & q & " " &_
    " -Y " & q & sLogFile & q & " " 	
  'WScript.Echo sCmd			' debug
  sResultado = RunAsAgente(sCmd)
  
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
  AddToLog="NAO"
  Const ForAppending = 8
  set objFile = oFS.OpenTextFile(sLogFile, ForAppending, True)
  sMensagem = Replace(sMensagem,sEscondeSenha1,"*****") 
  sMensagem = Replace(sMensagem,sEscondeSenha2,"*****") 
  objFile.WriteLine(DataExtenso(Now(), true) & ";" & sMensagem)
  If (Err.Number <> 0) and ( sSemParar<>"SIM" ) Then
    WScript.Echo "Adicionando mensagem ao arquivo :" & sLogFile & vbCrlf & _
      "C�digo do Erro: " & Err.Number & vbCrlf & _
      "C�digo do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
      "Fonte: " &  Err.Source & vbCrlf & _
      "Descri��o do Erro: " &  Err.Description
    Err.Clear
  else
    AddToLog="SIM"
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
  if sExibeHoras=True Then Resultado=Resultado & "+" & sHour & "h" & sMin & "m" & sSec & "s"
	
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

' Extrai a extens�o de um arquivo que esteja numa string
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
  RunAsAgente = "NAO"
  If Not oFS.FileExists(ACMD) Then
	  oWS.run ACMD, 1, True  
	  'oWS.run(ACMD)
	  If Err.Number <> 0 Then
		 sMensagem = vbTab & "  Falhou :" & ACMD & vbCrlf & _
		   "C�digo do Erro: " & Err.Number & vbCrlf & _
		   "C�digo do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
		   "Fonte: " &  Err.Source & vbCrlf & _
		   "Descri��o do Erro: " &  Err.Description
		 AddToLog(sMensagem)
		 if  ( sSemParar<>"SIM" ) Then WScript.Echo( sMensagem )
	  else
		 RunAsAgente="SIM"
		 sMensagem = vbTab & "Sucesso : " & ACMD
		 AddToLog(sMensagem)	   
	  End If    
	  Err.Clear
  End If
End Function

' Elimina arquivos de uma pasta que sejam mais antigos que a 
' data atual - iDaysOld 
Sub LimparArquivosBackupsAntigos(sDirectoryPath, iDaysOld)
Dim oFolder, oFileCollection, oFile, sPodeApagar, sFileName  
Dim sDir   
  Set oFS = CreateObject("Scripting.FileSystemObject") 
  Set oFolder = oFS.GetFolder(sDirectoryPath) 
  Set oFileCollection = oFolder.Files 

  For each oFile in oFileCollection
    'No exemplo abaixo, apenas arquivos com a extens�o .dat seriam apagados
    sPodeApagar="N" 
    sFileName=Cstr(sDirectoryPath) & "\" & Cstr(oFile.Name)
    If LCase(Right(Cstr(oFile.Name), 3)) = "fbk" Then sPodeApagar="S"
    If LCase(Right(Cstr(oFile.Name), 3)) = "dat" Then sPodeApagar="S"
    If LCase(Right(Cstr(oFile.Name), 3)) = "log" Then sPodeApagar="S"
    If sPodeApagar = "S" Then
      If oFile.DateLastModified < (Date() - iDaysOld) Then 
        oFile.Delete(True)   
        If Err.Number <> 0 Then
          sMensagem = vbTab & "  Falhou ao eliminar arquivo :" & sFileName & vbCrlf & _
            "C�digo do Erro: " & Err.Number & vbCrlf & _
            "C�digo do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
            "Fonte: " &  Err.Source & vbCrlf & _
            "Descri��o do Erro: " &  Err.Description
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
