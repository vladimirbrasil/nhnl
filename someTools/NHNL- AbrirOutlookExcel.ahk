;#include Macro- Funcoes.ahk	;inclui funções básicas minhas

;^0::
	;Inicialização
	imgOutlook := "OutExc_imgOutlook.gif"
	imgExcel := "OutExc_imgExcel.gif"
	imgOutMacro := "OutExc_imgOutMacro.gif"

	SetTitleMatchMode 2
	nTempoEspera:= 5
	
;		fPausaEBloqueia(10)
	Send, #d ;Minimiza tudo
		fPausaEBloqueia(3)

	;Abre Excel (e maximiza)
	Run, "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE"
	a := EsperarImagem(X, Y, imgExcel, 7000)

		fPausaEBloqueia()

;	WinMaximize, Microsoft Excel
;		fPausaEBloqueia()

	;Abre Outlook
	Run, "C:\Program Files (x86)\Microsoft Office\Office12\OUTLOOK.EXE"
	a := EsperarImagem(X, Y, imgOutlook, 10000)
		fPausaEBloqueia()

		
	AtivaJanela ("Microsoft Outlook")
		fPausaEBloqueia()

	;Abre janela de Macros
	Send, !{F8}
	AtivaJanela ("Microsoft Outlook")
	Send, !{F8}
	Send, !{F8}
		fPausaEBloqueia()
	Send, !{F8}
	a := EsperarImagem(X, Y, imgOutMacro, 7000)
		fPausaEBloqueia()

	;Seleciona primeira macro (NHNL_0)
	AtivaJanela ("Macros")
		fPausaEBloqueia()
	Send, {ENTER}
		fPausaEBloqueia()
	Send, {ENTER}

	;Encerramento
;	BlockInput, off
	SetTitleMatchMode 1

	sTrayTipFinal := "Agora já pode trocar de usuário ou usar o computador normalmente.`nNão feche o Outlook nem o Excel nos próximos trinta minutos.`nSwitch User ocorrerá em 30 segundos. Pressione ESC para cancelar."
	TrayTip, Obrigado!, %sTrayTipFinal%, 30, 0

	;Timer para logoff
	Period := 1 ;20
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

fPausaEBloqueia(i=1, sTrayTip="Não mexa no computador enquanto permanecer esta mensagem.")
{
	nPausa_1 := 500
	nPausa_2 := 4000
	nPausa_3 := 10000
	nPausa_4 := 25000

	TrayTip, Aguarde!, %sTrayTip%, 30, 2

;	if i = 0
;		;sem pausa
	if i = 1
		Sleep, nPausa_1
	else if i = 2
		Sleep, nPausa_2 
	else if i = 3
		Sleep, nPausa_3 
	else if i = 4
		Sleep, nPausa_4 
	else
		Sleep, nPausa_1 

	BlockInput, on
}

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

	;Bloquear ação do usuário
	BlockInput, on

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

EsperarImagem(ByRef X, ByRef Y, sLocalImagem, nPausa=10000, sJanela="")
{

	count := 0
	X := 0
	Y := 0
	nTolerancia := 20
	
TentarDeNovo_EsperarImagem:

	if sJanela <> ""
		AtivaJanela (sJanela)

	a := ProcurarImagem(X, Y, sLocalImagem, nTolerancia)
	if a = 0
	{
		Sleep, 200
		return ErrorLevel
	}
	else
	{
		if count < 25
		{
			count++
			
			;Aumentar tolerancia de tempos em tempos
			if count > 7 and count <= 12
				nTolerancia := 30
			if count > 12 and count <= 17
				nTolerancia := 40
			if count > 17
				nTolerancia := 60
				
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