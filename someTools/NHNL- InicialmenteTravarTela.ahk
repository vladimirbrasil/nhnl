#Persistent
#SingleInstance Force
#NoTrayIcon
;WinMinimizeAll

	nPausa := 4000
	nPausaMaior := 7000
	nPausaMuitoMaior := 15000
	nTempoEspera:= 5
	
	Period:=20 ;nPausaMaior+7*nPausa+2*nPausaMaior+4*nPausaMuitoMaior

Gui, +AlwaysOnTop -Disabled -SysMenu +Owner -Caption -ToolWindow
Gui, Font, s27 cFFFFFF, Ariel
pos_x := (A_ScreenWidth/2)-100
pos_y := (A_ScreenHeight/2)-40
pos_x2 := (A_ScreenWidth-45)
;Gui, Add, Button, x%pos_x% y%pos_y% h80 w200 gHibernate, Hibernate 
Gui, Add, Button, xs-250  y%pos_y% h80 w200 gLockPC, Lock PC
;Gui, Add, Button, xs-500  y%pos_y% h80 w200 gSwitchUser, Switch user
;Gui, Add, Button, xs+500  y%pos_y% h80 w200 gShutdown, Shutdown
;Gui, Add, Button, xs+250  y%pos_y% h80 w200 gRestart, Restart
;Gui, Add, Button, x%pos_x% ys+130 h80 w200 gCancel, Cancel
;Gui, Add, Button, x%pos_x2% y10 h30 w30 gCancel, X
Gui, Font, s55 cFFFFFF, Ariel
Gui, Add, Text, xs-500 ys-200 w1400 vPeriod, Não mexa no computador nos próximos %Period%` segundos
Gui, Font, s35 cFFFFFF, Ariel
Gui, Color, 000000                                    
Gui, Show, x0 y0 h%A_ScreenHeight% w%A_ScreenWidth%, ScreenMask
WinSet, Transparent, 200, ScreenMask
SetTimer, WaitTimer, 1000
BlockInput, on
Return

WaitTimer:
Period -= 1
GuiControl,,Period,Não mexa no computador nos próximos %Period%` segundos
If ( Period=0 ) {
	Period=10
	SetTimer, LockTimer, 1000
;ShutDown, 4+1+8
;BlockInput, On
;DllCall("ShowCursor", "Int", 0)
;Gui, Destroy
   }
Return

LockTimer:
BlockInput, off
Period -= 1
GuiControl,,Period,Obrigado por aguardar. Troca de usuário em %Period%` segundos
Gui, Add, Button, x%pos_x% ys+130 h80 w200 gCancel, Cancel
Gui, Add, Button, x%pos_x2% y10 h30 w30 gCancel, X
If ( Period=0 ) {
	Gosub, LockPC
   }
Return

LockPC:
;WinMinimizeAllUndo
Gui, Destroy
DllCall("LockWorkStation")
ExitApp
Return

Cancel:
;WinMinimizeAllUndo
Gui, Destroy
sleep 500
exitapp
return

Esc::
;WinMinimizeAllUndo
Gui, Destroy
sleep 500
;ExitApp