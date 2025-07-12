
	SetTitleMatchMode 2

	nPausaMinima := 500
	nPausa := 4000
	nPausaMaior := 10000
	nPausaMuitoMaior := 25000
	nTempoEspera:= 5

	BlockInput, on
	
	sTrayTip := "Não mexa no computador enquanto permanecer esta mensagem."
	sTrayTipFinal := "Agora já pode trocar de usuário ou usar o computador normalmente.`nNão feche o Outlook nem o Excel nos próximos trinta minutos.`nSwitch User ocorrerá em 30 segundos. Pressione ESC para cancelar."
	
	TrayTip, Aguarde!, %sTrayTip%, 30, 2

		Sleep, nPausa 
		BlockInput, on
	
	TrayTip, Aguarde!, %sTrayTip%, 30, 2

		Sleep, nPausaMuitoMaior 
		BlockInput, on
	
	TrayTip, Aguarde!, %sTrayTip%, 30, 2

		Sleep, nPausaMuitoMaior 
		BlockInput, on

	TrayTip, Aguarde!, %sTrayTip%, 30, 2

	Run, "C:\Program Files (x86)\Microsoft Office\Office12\OUTLOOK.EXE"
		Sleep, nPausaMaior 
		BlockInput, on

	TrayTip, Aguarde!, %sTrayTip%, 30, 2

	Run, "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE"
		Sleep, nPausaMaior 
		BlockInput, on

	TrayTip, Aguarde!, %sTrayTip%, 30, 2
		
	sNomeJanela := "Microsoft Excel"
	WinMaximize, %sNomeJanela%

		Sleep, nPausa 
		BlockInput, on

	sNomeJanela := "Microsoft Outlook"
	Gosub, AtivaJanela

	TrayTip, Aguarde!, %sTrayTip%, 30, 2

		Sleep, nPausaMuitoMaior 
		BlockInput, on
	
	sNomeJanela := "Microsoft Outlook"
	Gosub, AtivaJanela
		Sleep, nPausa 
		BlockInput, on
	sNomeJanela := "Microsoft Outlook"
	Gosub, AtivaJanela
		Sleep, nPausaMinima 
		BlockInput, on
	SendInput, !{F8}
;	Gosub, AtivaJanela
;		Sleep, nPausaMinima 
	SendInput, !{F8}
		Sleep, nPausaMinima 
		BlockInput, on
	SendInput, !{F8}
;	Gosub, AtivaJanela
	SendInput, !{F8}

		Sleep, nPausa 
		BlockInput, on

	TrayTip, Aguarde!, %sTrayTip%, 30, 2

		Sleep, nPausa 
		BlockInput, on

;	SendInput, NHNL
;
;		Sleep, nPausa 

		BlockInput, on
	sNomeJanela := "Macros"
	Gosub, AtivaJanela
		Sleep, nPausaMinima 
		BlockInput, on

	TrayTip, Aguarde!, %sTrayTip%, 30, 2	
		BlockInput, on
	
	SendInput, {ENTER}
		BlockInput, on

		Sleep, nPausa 
		BlockInput, off

;	TrayTip, Aguarde!, %sTrayTip%, 30, 2
		
	SetTitleMatchMode 1

;	sNomeJanela := "Microsoft Outlook"
;	WinMinimize, %sNomeJanela%

;		Sleep, nPausa 

;	TrayTip, Aguarde!, %sTrayTip%, 30, 2
		
;	sNomeJanela := "Microsoft Excel"
;	Gosub, AtivaJanela

;	WinMinimizeAll

;	BlockInput, off
	
;		Sleep, nPausa 

	TrayTip, Obrigado!, %sTrayTipFinal%, 30, 0

	Period := 30
	SetTimer, WaitTimer, 1000
	Return

WaitTimer:
	Period -= 1
	TrayTip, Obrigado!, Agora já pode trocar de usuário ou usar o computador normalmente.`nNão feche o Outlook nem o Excel nos próximos trinta minutos.`nSwitch User ocorrerá em %Period% segundos. Para cancelar clique com o botão direito no 'H' verde abaixo à direita e escolha 'Exit'., 30, 0
	;GuiControl,,Period,Agora já pode trocar de usuário ou usar o computador normalmente.`nNão feche o Outlook nem o Excel nos próximos trinta minutos.`nSwitch User ocorrerá em %Period% segundos. Pressione ESC para cancelar.
	If ( Period=0 ) {
		Gui, Destroy
		DllCall("LockWorkStation")
		ExitApp
		Return
	}
	Return

;		Sleep, nPausaMuitoMaior 
;		Sleep, nPausaMuitoMaior 

		
;SwitchUser: 
	;SwitchUser not working 
		;somente manualmente executa o arquivo C:\Windows\System32\tsdiscon.exe
		;pelo autohotkey dá erro. Usar LockPC então.
;	Run C:\Windows\System32\tsdiscon.exe
;	Process, priority, tsdiscon.exe, High
;	WinMinimizeAllUndo
;	Gui, Destroy
;	ExitApp

;LockPC:
;	DllCall("LockWorkStation")
;	ExitApp
;		
	return


AtivaJanela:
TentarDeNovo_AtivaJanela:

	WinWait, %sNomeJanela%, , %nTempoEspera%
	if ErrorLevel <> 0
	{
		;WinWait timed out. Não achou a janela.
		MsgBox, 4097, Mensagem da Macro, Abra o arquivo '%sNomeJanela%' e, em seguida, pressione OK.
		IfMsgBox, OK
			Goto, TentarDeNovo_AtivaJanela
		else 
			Exit
	}
	else
	{
		;Ação sobre a janela found by WinWait.
		IfWinExist, %sNomeJanela%
			WinActivate, %sNomeJanela% ;Ativa a janela encontrada acima
	}
return