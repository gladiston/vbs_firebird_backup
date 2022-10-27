# vbs_firebird_backup

Trata-se de um script vbs (Windows Script Host - VBScript) com o propósito de fazer backup de base de dados FirebirdSQL.
Use-o conjuntamente com o agendador de tarefas no servidor.
Ele faz o backup e deposita os arquivos numa pasta local ou remota.
<br>
Modo de usar:<br>
   vbs_firebird_backup.vbs \\\\server\bak "C:\\dados\\banco1.FDB" "C:\\dados\\banco2.FDB" "C:\\dados\\banco3.FDB"<br>
* Parametro #1: Refere-se ao destino do backup, pode ser uma pasta local ou pasta remota do tipo UNC como \\\\servidor\\compatilhamento. Se usar uma pasta remota junto com o agendador de tarefas, programe o agendador de tarefas para rodar a tarefa sob um usuario que tenha permissão a pasta remota e também que o enviroment deste usuário tenha as variaveis de ambiente ISC_USER e ISC_PASSWORD. Internamente este script faz o backup para a pasta temporaria do Windows, apenas depois de completo irá movê-lo para a unidade de destino que foi informada, esse método é mais ágil do que gerar o backup diretamente na unidade de destino principalmente porque não terá o lag da rede inicialmente e também porque vários computadores tem a unidade C: como um disco SSD/NVMe. Um alerta importante, se estiver num host que tenha serviços de sincronização com a nuvem como o onedrive, gdrive,... veja se o tempo de sincronização não compromete o backup, pois se acontecer um sinistro com o servidor antes que a sincronização termine e então terá um belo problema nas mãos.
* Parametro #2..N: Todos os parametros seguintes referem-se aos arquivos de dados que terão o seu backup realizado, mas atenção que devem estar em aspas duplas. Recomendo que caso opte por varios bancos de uma só vez que então coloque os bancos mais prioritarios primeiros. Caso use o agendador de tarefas do Windows com este script, crie uma programação onde os bancos mais importantes tenham um intervalo entre backups menor e os menos importantes com intervalos maiores, essa é a premissa de dividir para conquistar.<br>
* Este script usa as variaveis de ambiente ISC_USER e ISC_PASSWORD para saber qual usuario e senha que deverá ser usada para a realização do backup. Se quiser fazer diferente, voce pode modificar as referencias:<br>
   fdb_server="localhost"<br>
   fdb_username= oWS.ExpandEnvironmentStrings("%ISC_USER%")<br>
   fdb_password= oWS.ExpandEnvironmentStrings("%ISC_PASSWORD%")<br>
Porém ao modificá-las e colocar valores explicitos voce estaria sendo imprudente, pois se este script vazar, pessoas inescrupulosas poderiam usar essas informações explicitadas para invadir o seu sistema.<br>
* Este script tem suporte a voz(de bêbado) para indicar verbalmente quando inicia e quando termina, mas para ser sincero, não acho que isso seja importante especialmente em servidores, por isso caso queira desligar troque de True para False na linha:
  bWantVoice=True
