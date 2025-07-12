;pause::pause
;NumpadSub::

	SetTitleMatchMode 2
	CoordMode, Mouse, Screen

	nPausa := 600
	nPausaMaior := 2500
	nTempoEspera := 5

	MouseGetPos, Xuser, Yuser
		Sleep, nPausaMaior
		Sleep, nPausaMaior
		Sleep, nPausaMaior
	MouseMove, 0, 0
	
	sClipboard := clipboard
	
	sNomeJanela := "Microsoft Excel"
	Gosub, AtivaJanela
	
		Sleep, nPausa
		Sleep, nPausaMaior

;	Run, SnippingTool.exe
;snipping tool não roda no autohotkey
Run, "C:\Windows\Sysnative\SnippingTool.exe"

		Sleep, nPausaMaior
;	SendInput, ^n
		Sleep, nPausaMaior
		Sleep, nPausa

	sNomeJanela := "Snipping Tool"
	Gosub, AtivaJanela

		Sleep, nPausa
	Xo=127
	Yo=272 ;277 ;(Excel 2007)
	dx=1137
	dy=455
	MouseClickDrag, L, Xo,Yo,Xo+dx,Yo+dy

		Sleep, nPausa
	SendInput, ^s
		Sleep, nPausa
		Sleep, nPausa
;clipboard não funciona no autohotkey
;	SendInput, ^v

	SendInput, %sClipboard%
		Sleep, nPausa
		Sleep, nPausa
		Sleep, nPausa
		Sleep, nPausa		
	;Salvar
	SendInput, {ENTER}
		Sleep, nPausa
		Sleep, nPausa
	;Substituir existente? Sim.
	SendInput, y
		Sleep, nPausa
	;Garantir escapar de outras mensagens
	SendInput, {ESC}
		Sleep, nPausa

	sNomeJanela := "Snipping Tool"
	Gosub, AtivaJanela
		Sleep, nPausa
	WinClose
		Sleep, nPausa
	SendInput, !{TAB}

;	MouseMove, Xuser, Yuser ;Não voltar à região em cima do gráfico ativando a dica de texto

	
	CoordMode, Mouse, Relative	
	SetTitleMatchMode 1
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
		{
			WinActivate, %sNomeJanela% ;Ativa a janela encontrada acima
;			WinRestore, %sNomeJanela% ;Se estiver minimizada reaparece
		}
		else
		{
		}
	}
return