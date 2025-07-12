
	SetTitleMatchMode 1
	
	nPausa := 800
	nPausaMaior := 15000 ;Aguardar janela de segurança do outlook ficar habilitada
	nTempoEspera=5

		Sleep, nPausa
		Sleep, nPausaMaior
	
	sNomeJanela := "Microsoft Office Outlook" ;"Microsoft Outlook" ;(2010)

	IfWinExist, %sNomeJanela%
	{
		WinActivate

		SendInput, {SHIFTDOWN}{TAB}{SHIFTUP}
		Sleep, nPausa
		SendInput, {ENTER}
		Sleep, nPausa
	}
	
	SetTitleMatchMode 1
return
