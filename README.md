# vbs_firebird_backup

Trata-se de um script vbs (Windows Script Host - VBScript) com o propósito de fazer backup de base de dados FirebirdSQL.
Use-o conjuntamente com o agendador de tarefas no servidor.
Ele faz o backup e deposita os arquivos numa pasta local ou remota.
<br>
Modo de usar:<br>
   vbs_firebird_backup.vbs \\\\server\bak "C:\dados\banco1.FDB" "C:\dados\banco2.FDB" "C:\dados\banco3.FDB"<br>
* Parametro #1: Refere-se ao destino do backup, pode ser uma pasta local ou pasta remota do tipo UNC como \\servidor\comparilhamento. Se usar uma pasta remota junto com o agendador de tarefas, programe o agendador de tarefas para rodar a tarefa sob um usuario que tenha permissão a pasta remota. Um alerta importante, se estiver num host que tenha serviços de sincronização com a nuvem como o onedrive, gdrive,... estes serviços não sabem quando o backup terminou e por essa razão, quando o backup iniciar-se a cada byte de backup o programa irá querer sincronizar quando deveria esperar terminar primeira, isso não compromete o backup, mas tornará a sincronização para a nuvem bastante demorada.<br>
* Parametro #2...N: Todos os parametros seguintes referem-se aos arquivos de dados que terão o seu backup realizado, mas atenção que devem estasr em aspas duplas. Recomendo que caso opte por varios bancos de uma só vez que então coloque os bancos mais prioritarios primeiros. Caso use o agendador de tarefas do Windows, crie uma programação onde os bancos mais importantes tenham um intervalo entre backups menor e os menos importantes com intervalos maiores, essa é a premissa de dividir para 	conquistar.<br>
* Este script usa a variavel de ambiente ISC_USER e ISC_PASSWORD para saber qual usuario e senha que deverá ser usada para a realização do backup. Voce pode trocar as referencias:<br>
   fdb_server="localhost"<br>
   fdb_username= oWS.ExpandEnvironmentStrings("%ISC_USER%")<br>
   fdb_password= oWS.ExpandEnvironmentStrings("%ISC_PASSWORD%")    <br>
Ao modificar e colocar valores explicitos nessas variaveis voce estaria engessando o host, usuario e senha e não dependeria mais variaveis de ambiente, isso facilitaria, mas isto seria imprudente, pois se este script vazar, pessoas inescrupulosas poderiam usar essa informação para invadir o seu sistema.<br>

