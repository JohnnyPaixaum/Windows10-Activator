DIM KEYVER, TMPFILE, CmdShell, CmdScript
SET TMPFILE = CreateObject("Scripting.FileSystemObject")
SET CmdShell = CreateObject("Shell.Application")
Set CmdScript = WScript.CreateObject( "WScript.Shell" )

'Batch of serials of all versions:
KEY01="TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
KEY02="3KHY7-WNT83-DGQKR-F7HPR-844BM"
KEY03="PVMJN-6DFY6-9CCP6-7BKTT-D3WVR"
KEY04="7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"
KEY05="W269N-WFGWX-YVC9B-4J6C9-T83GX"
KEY06="MH37W-N47XK-V7XM9-C7227-GCQG9"
KEY07="6TP4R-GNPTD-KYYHQ-7B7DP-J447Y"
KEY08="YVWGF-BXNMC-HTQYQ-CPQ99-66QFC"
KEY09="NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"
KEY10="9FNHH-K3HBT-3W4TD-6383H-6XYWF"
KEY11="NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
KEY12="2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
KEY13="NPPR9-FWDCX-D2C8J-H872K-2YT43"
KEY14="DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
KEY15="YYVX9-NTFWV-6MDM3-9PT4T-4M68B"
KEY16="44RPN-FTY23-9VTTB-MP9BX-T84FV"
KEY17="WNMTR-4C88C-JK8YV-HQ7T2-76DF9"
KEY18="2F77B-TNFGY-69QQF-B8YKP-D69TJ"
KEY19="DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ"
KEY20="QFFDN-GRT3P-VKWWX-X7T3R-8B639"
KEY21="M7XTQ-FN8P6-TTKYV-9D4CC-J462D"
KEY22="92NFX-8DJQP-P6BBQ-THF9C-7CG2H"

WINVER = InputBox("SELECIONE UM NUMERO EQUIVALETE A SUA VERSAO DO WINDOWS 10: " &vbCRLF&_ 
"[01] Home/Core" &vbCRLF&_
"[02] Home/Core N" &vbCRLF&_	
"[03] Home/Core Country Specific" &vbCRLF&_
"[04] Home/Core Single Language" &vbCRLF&_	
"[05] Professional" &vbCRLF&_
"[06] Professional N" &vbCRLF&_
"[07] Professional Education" &vbCRLF&_	
"[08] Professional Education N" &vbCRLF&_	
"[09] Professional for Workstations" &vbCRLF&_
"[10] Professional for Workstations N" &vbCRLF&_	
"[11] Education" &vbCRLF&_
"[12] Education N" &vbCRLF&_	
"[13] Enterprise" &vbCRLF&_
"[14] Enterprise N" &vbCRLF&_
"[15] Enterprise G" &vbCRLF&_
"[16] Enterprise G N" &vbCRLF&_	
"[17] Enterprise 2015 LTSB"	&vbCRLF&_ 
"[18] Enterprise 2015 LTSB N" &vbCRLF&_
"[19] Enterprise 2016 LTSB" &vbCRLF&_	
"[20] Enterprise 2016 LTSB N" &vbCRLF&_	
"[21] Enterprise 2019 LTSC"	&vbCRLF&_
"[22] Enterprise 2019 LTSC N" &vbCRLF&_
"[99] QUAL A MINHA VERSAO?" &vbCRLF&_
"OBS:VERSOES COMUNS SAO 01,05 e 13","ATIVADOR PARA WINDOWS 10 V:BETA")

if WINVER = "01" then 
    KEYVER=KEY01
elseif WINVER = "02" Then
    KEYVER=KEY02
elseif WINVER = "03" Then
    KEYVER=KEY03
elseif WINVER = "04" Then
    KEYVER=KEY04
elseif WINVER = "05" Then
    KEYVER=KEY05
elseif WINVER = "06" Then
    KEYVER=KEY06
elseif WINVER = "07" Then
    KEYVER=KEY07
elseif WINVER = "08" Then
    KEYVER=KEY08
elseif WINVER = "09" Then
    KEYVER=KEY09
elseif WINVER = "10" Then
    KEYVER=KEY10
elseif WINVER = "11" Then
    KEYVER=KEY11
elseif WINVER = "12" Then
    KEYVER=KEY12
elseif WINVER = "13" Then
    KEYVER=KEY13
elseif WINVER = "14" Then
    KEYVER=KEY14
elseif WINVER = "15" Then
    KEYVER=KEY15
elseif WINVER = "16" Then
    KEYVER=KEY16
elseif WINVER = "17" Then
    KEYVER=KEY17
elseif WINVER = "18" Then
    KEYVER=KEY18
elseif WINVER = "19" Then
    KEYVER=KEY19
elseif WINVER = "20" Then
    KEYVER=KEY20
elseif WINVER = "21" Then
    KEYVER=KEY21
elseif WINVER = "22" Then
    KEYVER=KEY22
end if

If IsEmpty(WINVER) Then
    'cancelled
    MsgBox "A OPERACAO CANCELADA",48,"ATIVADOR PARA WINDOWS 10 V:BETA "
Elseif WINVER = 99 Then
    CmdShell.ShellExecute "winver.exe"
    WScript.Echo ("*****A VERSAO PODE ESTA NO SEGUNDO PARAGRAFO, REABRA O ATIVADOR!*****")
Else
    'something has entered even zero-length
    TMPFILE.CreateTextFile("C:\Users\Public\Documents\cache_ativador.temp").WriteLine(KEYVER)

    CmdScript.Popup "*****INICIANDO ATIVACAO!*****",, "ACTIVATOR FOR WINDOWS 10"

    CmdShell.ShellExecute "powershell.exe", "-noexit -Command " &_
            "ECHO *================================*; " &_
            "ECHO ***********DESENVOLVEDOR**********; " &_
            "ECHO *****GIT/INSTA:@johnnypaixaum*****; " &_
            "ECHO *================================*; " &_
        "$KEYVER = (cat C:\Users\Public\Documents\cache_ativador.temp); " &_
            "ECHO *==============*; " &_
            "ECHO *****STAP_1*****; " &_
            "ECHO *==============*; " &_
        "cscript.exe C:\Windows\System32\slmgr.vbs /ipk $KEYVER; " &_
            "ECHO *==============*; " &_
            "ECHO *****STAP_2*****; " &_
            "ECHO *==============*; " &_
        "cscript.exe C:\Windows\System32\slmgr.vbs /skms kms8.msguides.com; " &_
            "ECHO *==============*; " &_
            "ECHO *****STAP_3*****; " &_
            "ECHO *==============*; " &_
            "ECHO *=========================================*; " &_
            "ECHO *****ESPERE_A_FINALIZACAO_DO_PROCESSO!*****; " &_
            "ECHO *=========================================*; " &_
        "cscript.exe C:\Windows\System32\slmgr.vbs /ato; " &_
            "ECHO *===============================*; " &_
            "ECHO *****ATIVADO_COM_SUCESSO!*****; " &_
            "ECHO *===============================*; " &_
        "timeout 15; exit """, "", "runas", 1

WScript.Quit
WScript.Sleep(10000)
TMPFILE.DeleteFile("C:\Users\Public\Documents\cache_ativador.temp")
End If