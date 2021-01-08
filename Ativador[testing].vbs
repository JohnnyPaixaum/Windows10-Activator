REM ***********************
REM *      GIT&INSTA      *
REM *   @johnny.paixaum   *
REM ***********************

KEYS = Array("NUMERO INVARIDO!", _
"TX9XD-98N7V-6WMQ6-BX7FG-H8Q99", _
"3KHY7-WNT83-DGQKR-F7HPR-844BM", _
"PVMJN-6DFY6-9CCP6-7BKTT-D3WVR", _ 
"7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH", _
"W269N-WFGWX-YVC9B-4J6C9-T83GX", _
"MH37W-N47XK-V7XM9-C7227-GCQG9", _
"6TP4R-GNPTD-KYYHQ-7B7DP-J447Y", _
"YVWGF-BXNMC-HTQYQ-CPQ99-66QFC", _
"NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J", _
"9FNHH-K3HBT-3W4TD-6383H-6XYWF", _
"NW6C2-QMPVW-D7KKK-3GKT6-VCFB2", _
"2WH4N-8QGBV-H22JP-CT43Q-MDWWJ", _
"NPPR9-FWDCX-D2C8J-H872K-2YT43", _
"DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4", _
"YYVX9-NTFWV-6MDM3-9PT4T-4M68B", _
"44RPN-FTY23-9VTTB-MP9BX-T84FV", _
"WNMTR-4C88C-JK8YV-HQ7T2-76DF9", _
"2F77B-TNFGY-69QQF-B8YKP-D69TJ", _
"DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ", _
"QFFDN-GRT3P-VKWWX-X7T3R-8B639", _
"M7XTQ-FN8P6-TTKYV-9D4CC-J462D", _
"92NFX-8DJQP-P6BBQ-THF9C-7CG2H")

IN_WINVER = InputBox("SELECIONE UM NUMERO EQUIVALETE A SUA VERSAO DO WINDOWS 10: " &vbCRLF&_ 
"[99] QUAL A MINHA VERSAO?" &vbCRLF&_
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
"OBS:VERSOES COMUNS SAO 01,05 e 13","ATIVADOR PARA WINDOWS 10 v:latest")

If IsEmpty(IN_WINVER) Then
    'cancelled
    MsgBox "A OPERACAO CANCELADA",48,"ATIVADOR PARA WINDOWS 10 v:latest"
ELSEIf IN_WINVER = 99 Then
    'help version
    DIM GetSysExec, SYSINFO, objShell
    Set objShell = WScript.CreateObject("WScript.Shell")
    Set GetSysExec = objShell.Exec("powershell.exe -window hidden -Command " &_
    "$string01 = systeminfo.exe | findstr /B /C:'Nome do sistema operacional:'; $string01.TrimStart('Nome do sistema operacional:'); $SYSINFO ")
    SYSINFO = GetSysExec.StdOut.ReadLine()
    MsgBox "===========SUA VERSAO===========" &vbCRLF& SYSINFO &vbCRLF& "==============================",0+64,"ATIVADOR PARA WINDOWS 10 v:beta"      
ELSE 
    DIM TMPFILE, CmdAtv, CmdTool, KEYVER
    KEYVER = KEYS(IN_WINVER)
    SET TMPFILE = CreateObject("Scripting.FileSystemObject")
    SET CmdAtv = CreateObject("Shell.Application")
    Set CmdTool = WScript.CreateObject( "WScript.Shell" )
    TMPFILE.CreateTextFile("C:\Users\Public\Documents\cache_ativador.temp").WriteLine(KEYVER)
    CmdTool.Popup "*****INICIANDO ATIVACAO!*****",, "ACTIVATOR FOR WINDOWS 10 v:latest"
    CmdAtv.ShellExecute "powershell.exe", "-noexit -Command " &_
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
    WScript.Sleep(10000)
    TMPFILE.DeleteFile("C:\Users\Public\Documents\cache_ativador.temp")
End IF
