
$filtro = "Ultimo Acesso"

$g = @( -- , -- ) #unidade organizacional


#Esse array substitui o cabeçalho da tabela, referente ao select dado no laço
$nome = @{
     Label="Nome";
     Expression= {$_.name}
}
$login = @{
     Label="Login";
     Expression= {$_.samaccountname}
}
$acesso = @{
     Label="Ultimo Acesso";
     Expression= {$_.lastlogondate}
}

$param = $nome, $login, $acesso

$cliente = " ---  " #cliente onde está sendo executado o script

#laço que pega os usuarios dentro de uma OU, pode-se alterar a posição da variavel em cada ambiente; após isso o filtro
#pega todos os usuarios ATIVOS e seleciona as propriedades nome, nome de conta, e ultimo logon efetuado.
foreach ($Item in $g) {

     $sa += Get-ADUser -SearchBase "ou=$Item,dc= -- ,dc=com, dc=br" -filter {enabled -eq $true} -Properties *  |
      select $param 

     }

$data = Get-Date
#onde a pasta do CSS deve estar alocada para formatar a tabela
$estilo = Get-Content styles.css
#inserção do estilo da pagina
$styleTag = "<style> $estilo </style>"
#titulo da nossa tabela
$titulo = "Usuarios Ativos no AD $cliente Atualizado em $data"
#mostra o titulo
$titulobody = "<h1> $titulo </h1>"


#pega todos os usuarios e converte para HTML formatando, e em seguida a saida do arquivo e seu diretório
$sa | sort -p $filtro | ConvertTo-Html -Head $styleTag  -Title $titulo -Body $titulobody| Out-File $cliente.html 
$sa | sort -p $filtro | ConvertTo-Csv <#-Head $styleTag  -Title $titulo -Body $titulobody#>| Out-File $cliente.csv 
#limpa a saida para não somar usurios no AD
$sa.count 
$sa = $null
