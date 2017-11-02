<#
 #  relatorio_rdp.ps1
 #
 #	Relatorio de Conexoes via RDP
 #	Extrai o relatorio de todas as conexoes realizada em um ou mais servidores
 #
 #	Por: Marcos Henrique | www.100security.com.br
 #
 # 
#>


$hosts = @(
			'SRV-2008'
			
#			'HOST01',
#			'HOST02',
#			'HOST03',
#			'HOST04'

			)

foreach ($servidor in $hosts) {

    $LogFilter = @{
        LogName = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
        ID = 21, 23, 24, 25
        }

    $entradas = Get-WinEvent -FilterHashtable $LogFilter -ComputerName $servidor

    $entradas | Foreach { 
           $entrada = [xml]$_.ToXml()
        [array]$saida += New-Object PSObject -Property @{
            DATA_HORA = $_.TimeCreated
            USUARIO = $entrada.Event.UserData.EventXML.User
            COMPUTADOR = $entrada.Event.UserData.EventXML.Address
            EventID = $entrada.Event.System.EventID
            HOST = $servidor
            }        
           } 

}

$exportar += $saida | Select DATA_HORA, USUARIO, HOST, COMPUTADOR, @{Name='STATUS';Expression={
            if ($_.EventID -eq '21'){"LOGON"}
            if ($_.EventID -eq '22'){"SHELL START"}
            if ($_.EventID -eq '23'){"LOGOFF"}
            if ($_.EventID -eq '24'){"DESCONECTADO"}
            if ($_.EventID -eq '25'){"RECONECTADO"}
            }
        }

$data = (Get-Date -Format d) -replace "/", "-"

# Formatacao HTML

$a = "<style>"
$a = $a + "BODY{font-family: Calibri, Arial, Helvetica, sans-serif;font-size:10;font-color: #000000}"
$a = $a + "TABLE{margin-left:auto;margin-right:auto;width: 800px;border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color: #F9F9F9;text-align:center;}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align:center;}"
$a = $a + "</style>"

# Exportando para HTML:

$b = "<br><center><img src='http://www.100security.com.br/100security.png'><br><font face='Calibri' size='5'>Relatório de Conexões via RDP</font></center><br>"

$exportar | Sort "DATA_HORA" –des | ConvertTo-Html -head $a -body $b | Set-Content Relatorio_RDP_$data.html
