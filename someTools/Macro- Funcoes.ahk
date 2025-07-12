Pause::Pause ; Press {PAUSE} to pause. Press it again to resume.

; :::::
; Ativa a janela especificada
; ::::: 

AtivaJanela(sNomeJanela, nTempoEspera=5)
{

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
	return ErrorLevel   ; "Return" expects an expression.
}

; :::::
; ::::: 
; ::::: 


; :::::
; Pega conteúdo na clipboard
; ::::: 

PegaClipboard(sDadoParaCopiar="", nTempoEspera=5, sCopiaAntes=False)
{
nTentativas := 0

TentarDeNovo_PegaClipboard:

	nTentativas := nTentativas + 1

	if sCopiaAntes = True
	{
		if (nTentativas = 1)
			Clipboard = ;Limpa a clipboard
		Send ^c
	}

	ClipWait, nTempoEspera
	if ErrorLevel <> 0
	{
;		MsgBox, 4097, Mensagem da Macro, Selecione e copie (CTRL+C) '%sDadoParaCopiar%' e, em seguida, clique OK.
		If (nTentativas < 5)
			Goto, TentarDeNovo_PegaClipboard
		else 
			Exit
	}
	sClipSaved = %Clipboard%
	Clipboard = ;Limpa a Clipboard
	return sClipSaved
}

; :::::
; ::::: 
; ::::: 

; :::::
; Pega conteúdo na clipboard
; ::::: 

PegaClipboard_Tentativas(nTempoEspera=5)
{
nTentativas := 0

TentarDeNovo_PegaClipboard_Tentativas:

	nTentativas := nTentativas + 1

	if (nTentativas = 1)
		Clipboard = ;Limpa a clipboard
	Send ^c

	ClipWait, nTempoEspera
	if ErrorLevel <> 0
	{
		If (nTentativas < 5)
			Goto, TentarDeNovo_PegaClipboard_Tentativas
		else 
			Exit
	}
	sClipSaved = %Clipboard%
	Clipboard = ;Limpa a Clipboard
	return sClipSaved
}

; :::::
; ::::: 
; ::::: 

; :::::
; Espera conteúdo na clipboard
; ::::: 

EsperaClipboard(sOutroPrograma="", nTempoEspera=30)
{

	Clipboard = ;Limpa a clipboard

TentarDeNovo_EsperaClipboard:

	ClipWait, nTempoEspera

	if ErrorLevel <> 0
	{
		MsgBox, 4097, Mensagem da Macro, O programa '%sOutroPrograma%' não respondeu. Aguardar resposta?
		IfMsgBox, OK
			Goto, TentarDeNovo_EsperaClipboard
		else 
			Exit
	}
	sClipSaved = %Clipboard%
	sClipSaved := EliminarEspacos(sClipSaved)
	
	Clipboard = ;Limpa a Clipboard

	return sClipSaved
}

; :::::
; ::::: 
; ::::: 

; :::::
; Aguarda Janela Fechar
; ::::: 

AguardaJanelaFechar(sNomeJanela, nTempoEspera=10)
{

TentarDeNovo_AguardaJanelaFechar:

	WinWait, %sNomeJanela%, , 5
	if ErrorLevel <> 0
	{
		;WinWait timed out. Não achou a janela.
		MsgBox, 4097, Mensagem da Macro, Abra o arquivo '%sNomeJanela%' e, em seguida, pressione OK.
		IfMsgBox, OK
			Goto, TentarDeNovo_AguardaJanelaFechar
		else 
			Exit
	}
	else
	{
		;Ação sobre a janela found by WinWait.
		IfWinExist, %sNomeJanela%
		{
			WinWaitClose, %sNomeJanela%,, %nTempoEspera% ;Wait for the exact window found by WinWait to be closed.
			if ErrorLevel <> 0
			{
				;WinWaitClose timed out. A janela não fechou.
				MsgBox, 4097, Mensagem da Macro, A janela '%sNomeJanela%' não fechou. Após fechar a janela '%sNomeJanela%', pressione OK para prosseguir a macro.
				IfMsgBox, OK
				{
				}
				else 
					Exit
				
			}
		}
		else
		{
		}
	}
	return ErrorLevel   ; "Return" expects an expression.
}

; :::::
; ::::: 
; ::::: 

; :::::
; Fecha a janela especificada (se for encontrada).
; ::::: 

FechaJanela(sNomeJanela, nTempoEspera=2)
{

	WinWait, %sNomeJanela%, , %nTempoEspera%
	if ErrorLevel <> 0
	{
		;WinWait timed out. Não achou a janela.
	}
	else
	{
		;Ação sobre a janela found by WinWait.
		IfWinExist, %sNomeJanela%
		{
			WinClose, %sNomeJanela%
		}
		else
		{
		}
	}
	return ErrorLevel   ; "Return" expects an expression.
}

; :::::
; ::::: 
; ::::: 


; :::::
; Pesquisar nome fornecido.
; ::::: 

PesquisarNome(sNome)
{

	 ;Pesquisar nome completo agora
	sNome = sNome

}

; :::::
; ::::: 
; ::::: 

; :::::
; Procura imagem fornecida. Retorna posição X, Y para clique nela, por exemplo.
; Trabalhar com imagens pequenas, do tamanho de ícones pequenos foi testado com sucesso.
; Ocorre, em algumas tentativas mais, um aumento automático da tolerância em caso de não encontrar o ícone.
; ::::: 

ProcurarImagem(ByRef X, ByRef Y, sLocalImagem, nTolerancia=20)
{
	count := 0
	X := 0
	Y := 0

TentarDeNovo_ProcurarImagem:

	ImageSearch, X, Y, 0, 0, A_ScreenWidth, A_ScreenHeight, *%nTolerancia% %sLocalImagem%
	if ErrorLevel = 2
		MsgBox Problema na pesquisa.
	else if ErrorLevel = 1
		if count < 2
		{
			count++
			nTolerancia += 20
			Goto, TentarDeNovo_ProcurarImagem 
		}
		else
		{
;			MsgBox, 4097, Mensagem da Macro, Imagem não foi encontrada.
		}
	else
	{
		X += 10
		Y += 10
	}
	return ErrorLevel
}

; :::::
; ::::: 
; ::::: 

; :::::
; Espera imagem fornecida. Retorna posição X, Y para clique nela, por exemplo.
; Trabalhar com imagens pequenas, do tamanho de ícones pequenos foi testado com sucesso.
; Ocorre, em algumas tentativas mais, um aumento automático da tolerância em caso de não encontrar o ícone.
; ::::: 

EsperarImagem(ByRef X, ByRef Y, sLocalImagem, nPausa=20000)
{

	count := 0
	X := 0
	Y := 0

TentarDeNovo_EsperarImagem:

	a := ProcurarImagem(X, Y, sLocalImagem, 0)
	if a = 0
	{
		Sleep, 200
		return ErrorLevel
	}
	else
	{
		if count < 1000
		{
			count++
			Sleep, %nPausa%
			Goto, TentarDeNovo_EsperarImagem 
		}
		else
			MsgBox, 4097, Mensagem da Macro, Imagem não foi encontrada. A internet pode estar lenta.
	}
	return ErrorLevel
}

; :::::
; ::::: 
; ::::: 



; :::::
; Elimina espaços a mais
; ::::: 

EliminarEspacos(sNome)
{

	 ;Eliminar espaços das pontas (AutoTrim On por default)
	sTemp := sNome
	sNome = %sTemp%

	 ;Eliminar espaços internos
	StringReplace, sNome, sNome, %A_Space%%A_Space%,%A_Space%, All	

	return SNome

}

; :::::
; ::::: 
; ::::: 

; :::::
; Elimina preposições
; ::::: 

EliminarPreposicoes(sNome)
{

	 ;Eliminar preposições
	StringReplace, sNome, sNome, da%A_Space%,, All	
	StringReplace, sNome, sNome, do%A_Space%,, All	
	StringReplace, sNome, sNome, das%A_Space%,, All	
	StringReplace, sNome, sNome, dos%A_Space%,, All	
	StringReplace, sNome, sNome, de%A_Space%,, All	
	StringReplace, sNome, sNome, del%A_Space%,, All	
	StringReplace, sNome, sNome, dal%A_Space%,, All	

	return SNome

}

; :::::
; ::::: 
; ::::: 

