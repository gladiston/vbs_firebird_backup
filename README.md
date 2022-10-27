# vbs_firebird_backup

 Windows Script Host - VBScript
-----------------------------------------------------------------
 Nome: vbs_firebird_backup.vbs
 Proposito : Realização de Backup do Servidor FB
 By: Gladiston Santana (gladiston.santana em gmail.com)
 Copyright: (c) Jun 2011, Todos os direitos reservados!
-----------------------------------------------------------------

Modo de usar:
 vbs_firebird_backup.vbs \\server\bak "C:\dados\banco1.FDB" "C:\dados\banco2.FDB" "C:\dados\banco3.FDB" "C:\dados\banco4.FDB" "C:\dados\banco5.FDB" 
  Parametro #1: Refere-se ao destino do backup, pode ser uma pasta local ou 
    pasta remota UNC como \\servidor\comparilhamento. Se estiver num host que 
    tenha serviços de sincronização com a nuvem como o onedrive, gdrive,...
    tornará a sincronização bastante demorada e isso pode ser um risco para
    sua segurança.
  Parametro #2...N: Todos os parametros seguintes referem-se aos arquivos de 
    dados que terão o seu backup realizado, mas atenção que devem estasr em 
    aspas duplas. Recomendo que caso opte por varios bancos de uma só vez
	que então coloque os bancos mais prioritarios primeiros. Caso use o
	agendador de tarefas do Windows, crie uma programação onde os bancos
	mais importantes tenham um intervalo entre backups menor e os menos 
	importantes com intervalos maiores, essa é a premissa de dividir para
	conquistar.
  Este script usa a variavel de ambiente ISC_USER e ISC_PASSWORD para saber 
    qual usuario e senha que deverá ser usada para a realização do backup.
	 Voce pode trocar as referencias:
      fdb_server="localhost"
      fdb_username= oWS.ExpandEnvironmentStrings("%ISC_USER%")
      fdb_password= oWS.ExpandEnvironmentStrings("%ISC_PASSWORD%")    
    Ao modificar e colocar valores explicitos nessas variaveis voce estaria
	  engessando o host, usuario e senha e não dependeria mais variaveis de
	  ambiente, isso facilitaria, mas isto seria imprudente, pois se este 
	  script vazar, pessoas inescrupulosas poderiam usar essa informação 
	  para invadir o seu sistema.

