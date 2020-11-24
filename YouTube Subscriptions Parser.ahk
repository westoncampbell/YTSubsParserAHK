; Script Information ===========================================================
; Name:        YouTube Subscriptions Parser
; Description: Extracts a complete list of subscriptions from a YouTube account
; AHK Version: 1.1.33.02 (Unicode 32-bit, Unicode 64-bit, ANSI 32-bit)
; OS Version:  Windows 10
; Language:    English (United States)
; Author:      TheDewd (Weston Campbell <westoncampbell@gmail.com>)
;              autohotkey.com/boards/memberlist.php?mode=viewprofile&u=56166
; Filename:    YouTube Subscriptions Parser.ahk
; ==============================================================================

; Release History ==============================================================
; YYYY.MM.DD.Build.Release
; * TBD
; ------------------------------------------------------------------------------
; 2020.11.23.1.1
; * Initial release
; ==============================================================================

; Auto-Execute =================================================================
#SingleInstance, Force ; Allow only one running instance of script
#Persistent ; Keep the script permanently running until terminated
#NoEnv ; Avoid checking empty variables for environment variables
#Warn ; Enable warnings to assist with detecting common errors
#NoTrayIcon ; Disable the tray icon of the script
;#KeyHistory, 0 ; Keystroke and mouse click history
;ListLines, Off ; The script lines most recently executed
SetWorkingDir, % A_ScriptDir ; Set the working directory of the script
SetBatchLines, -1 ; The speed at which the lines of the script are executed
SendMode, Input ; The method for sending keystrokes and mouse clicks
DetectHiddenWindows, On ; The visibility of hidden windows by the script
;SetWinDelay, 0 ; The delay to occur after modifying a window
;SetControlDelay, 0 ; The delay to occur after modifying a control
;CoordMode, Menu, Window
OnExit("OnUnload") ; Run a subroutine or function when exiting the script

return ; End automatic execution
; ==============================================================================

; Functions ====================================================================
OnLoad() {
	Global ; Assume-global mode
	Static Init := OnLoad() ; Call function

	YouTube16 := "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAd0lEQVR4AWNwL/ChCFPPgP8MDAJA7ADFDVgxQl4Apg+mOQGI/5OIE8AGABkGIAEysQHIgAAKDEgAGdCAU8GDB///OzjgM6ABrwFwsGHD//8KChQYcOAAyQYgvBAQgNcL+ANRQIBgIFIcjRQlJHxJuYDYpDzwuREAvhinoHRyxq8AAAAASUVORK5CYII="

	YouTube32 := "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAuklEQVR4Ae2WBxWDMBRFI+FLiAQk/LO6lwQkVEIkIAUJkYCUOHhlde/52pJ/zmXDu6wkpr+cUIkCUeDkRhgjJVqSlrgWfyfr89L2WnJJYB2ctCfjTfiS5JTAOjyU4M2EtcShQF6CD5GfEsAnWeeuw20JPozsCuhdJy8WrxDQxwWq8h6wliSwrhAA54gC6yoKQJUosK4sA0S6JUB4BYSPkPAbEhoiblNM74z43TF/QMIfkvEHpSyiQBRYAcMdhq81rUIIAAAAAElFTkSuQmCC"

	hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll") ; Load module
	VarSetCapacity(GdiplusStartupInput, (A_PtrSize = 8 ? 24 : 16), 0) ; GdiplusStartupInput structure
	NumPut(1, GdiplusStartupInput, 0, "UInt") ; GdiplusVersion
	VarSetCapacity(pToken, 0)
	DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &GdiplusStartupInput, "Ptr", 0) ; Initialize GDI+

	pIcon16 := GdipCreateFromBase64(YouTube16) ; 16x16
	pIcon32 := GdipCreateFromBase64(YouTube32) ; 32x32

	Menu, Tray, Icon, % "HICON:*" pIcon16 ; Tray icon
	Menu, Tray, Tip, YouTube Subscriptions Extract

	KEY_FBE := "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
	EXE_NAME := (A_IsCompiled ? A_ScriptName : StrSplit(A_AhkPath, "\").Pop())

	RegRead, FBE, % KEY_FBE, % EXE_NAME
	RegWrite, REG_DWORD, % KEY_FBE, % EXE_NAME, 0
}

OnUnload(ExitReason, ExitCode) {
	Global ; Assume-global mode

	If (FBE = "") {
		RegDelete, % KEY_FBE, % EXE_NAME
	} Else {
		RegWrite, REG_DWORD, % KEY_FBE, % EXE_NAME, % FBE
	}

	DllCall("User32.dll\DestroyIcon", "Ptr", pIcon16)
	DllCall("User32.dll\DestroyIcon", "Ptr", pIcon32)
	DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
	DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
}

WM_RBUTTONDOWN(wParam, lParam, Msg, Hwnd) {
	Global ; Assume-global mode
	Static Init := OnMessage(0x0204, "WM_RBUTTONDOWN") ; Call function

	If (A_Gui = "YouTube") {
		return 0
	}
}

WM_RBUTTONDBLCLK(wParam, lParam, Msg, Hwnd) {
	Global ; Assume-global mode
	Static Init := OnMessage(0x0206, "WM_RBUTTONDBLCLK") ; Call function

	If (A_Gui = "YouTube") {
		return 0
	}
}

WM_KEYDOWN(wParam, lParam, Msg, Hwnd) {
	Global ; Assume-global mode
	Static Init := OnMessage(0x0100, "WM_KEYDOWN") ; Call function

	If (Chr(wParam) ~= "(?![ACVX])[A-Z]" || wParam = 0x74) {
		return ; Disable Ctrl+O/L/F/N and F5.
	}

	; IOleInPlaceActiveObject Interface
	IPAO := ComObjQuery(WB, "{00000117-0000-0000-C000-000000000046}")

	VarSetCapacity(kMsg, (A_PtrSize = 4 ? 24 : 48))
	NumPut(hWnd, kMsg)
	NumPut(Msg, kMsg, (A_PtrSize = 4 ? 4 : 8), "UInt")
	NumPut(wParam, kMsg, (A_PtrSize = 4 ? 8 : 16), "Int")
	NumPut(lParam, kMsg, (A_PtrSize = 4 ? 12 : 24), "Int")
	NumPut(A_EventInfo, kMsg, (A_PtrSize = 4 ? 16 : 32))
	NumPut(A_GuiX, kMsg, (A_PtrSize = 4 ? 20 : 40))
	NumPut(A_GuiY, kMsg, (A_PtrSize = 4 ? 24 : 48))

	Loop, 2 {
		r := NumGet(1 * IPAO)
		r := NumGet(r + (5 * A_PtrSize))
		r := DllCall(r, "Ptr", IPAO, "Ptr", &kMsg)
	} Until (wParam <> 9 || WB.document.activeElement <> "")

	ObjRelease(IPAO)

	If (r = 0) {
		return 0
	}
}

GuiCreate() {
	Global ; Assume-global mode
	Static Init := GuiCreate() ; Call function

	DetectHiddenWindows, On

	; App ----------------------------------------------------------------------
	Menu, SBUserMenu, Add, &Logout, MenuHandler
	Menu, SBUserMenu, Add ; Separator
	Menu, SBUserMenu, Add, &Switch User, MenuHandler
	Menu, FileMenu, Add, E&xit`tAlt+F4, MenuHandler
	Menu, HelpMenu, Add, &About...`tF1, MenuHandler
	Menu, MenuBar, Add, &File, :FileMenu
	Menu, MenuBar, Add, &Help, :HelpMenu
	Gui, App: Menu, MenuBar
	Gui, App: +LastFound +Resize +HWNDhApp +MinSize420x280 +OwnDialogs
	Gui, App: Default
	Gui, App: Margin, 10, 10
	Gui, App: Font, s10, Consolas
	Gui, App: Add, Edit, xm ym w0 h0 +ReadOnly ; Catch Focus
	Gui, App: Add, Edit, xm ym w340 r20 vEdt +HWNDhEdt +ReadOnly -Wrap +HScroll +VScroll, <No Data>
	Gui, App: Font
	Gui, App: Add, Button, w80 h24 vBtnStart +HWNDhBtnStart gControlHandler +Default, &Start
	Gui, App: Add, StatusBar, vSB +HWNDhSB gControlHandler +0x800, Idle
	Gui, App: Show, w600 h400, YouTube Subscriptions Parser
	GuiControl, Focus, % hBtnStart
	SB_SetParts(300)
	SBSetText("Idle", "YouTube User")
	SendMessage, 0x80, 0, pIcon16,, ahk_id %hApp% ; ICON_SMALL (16x16)
	SendMessage, 0x80, 1, pIcon32,, ahk_id %hApp% ; ICON_BIG (32x32)

	; About --------------------------------------------------------------------
	AboutText =
	(LTrim
		YouTube Subscriptions Parser
		2020.11.23.1.1

		Developed for AutoHotkey by TheDewd (Weston Campbell <westoncampbell@gmail.com>).
	)

	Gui, About: +LastFound +HWNDhAbout -Resize -MinimizeBox +OwnerApp +OwnDialogs
	Gui, About: Default
	Gui, About: Margin, 10, 10
	Gui, About: Font, s10, Consolas
	Gui, About: Add, Edit, xm ym w0 h0 +ReadOnly ; Catch Focus
	Gui, About: Add, Edit, xm ym w400 r12 vAboutEdt +HWNDhAboutEdt +ReadOnly, % AboutText
	Gui, About: Font
	Gui, About: Add, Button, x100 y+10 w80 h24 gControlHandler vBtnAboutOK HWNDhBtnAboutOK +Default, &OK
	Gui, About: Add, Button, xm yp w80 h24 gControlHandler vBtnAboutAHK HWNDhBtnAboutAHK, &AutoHotkey
	Gui, About: Add, Button, x+10 yp w80 h24 gControlHandler vBtnAboutGitHub HWNDhBtnAboutGitHub, &GitHub
	Gui, About: Show, AutoSize Hide, About
	GuiControlGet, AboutEdt, Pos
	GuiControl, MoveDraw, % hBtnAboutOK, % "x" (AboutEdtX + AboutEdtW) - 80

	; YouTube ------------------------------------------------------------------
	Gui, YouTube: +LastFound +Resize -DPIScale +HWNDhYouTube +OwnerApp +OwnDialogs
	Gui, YouTube: Default
	Gui, YouTube: Margin, 0, 0
	Gui, YouTube: Color, FFFFFF
	Gui, YouTube: Add, ActiveX, x0 y0 w640 h480 vWB +HWNDhWB, about:blank
	Gui, YouTube: Show, w400 h460 Hide, YouTube
	SendMessage, 0x80, 0, pIcon16,, ahk_id %hYouTube% ; ICON_SMALL (16x16)
	SendMessage, 0x80, 1, pIcon32,, ahk_id %hYouTube% ; ICON_BIG (32x32)

	WB.silent := True
	ComObjConnect(WB, WB_Events)
}

YT_GetData() {
	Global ; Assume-global mode

	GuiControl, Disable, % hBtnStart
	SBSetText("Retrieving data...")

	WB.Document.location.href := "https://m.youtube.com/feed/channels?app=m&persist_app=1"

	While (!InStr(WB.Document.location.href, "app=m"))
	&& (!InStr(WB.Document.location.href, "accounts.google")) {
		Sleep, 250
	}

	If (InStr(WB.Document.location.href, "accounts.google")) {
		If (!DllCall("User32.dll\IsWindowVisible", "UInt", WinExist("ahk_id " hYouTube))) {
			SBSetText("Waiting for login...")
			Gui, App: +Disabled
			Gui, YouTube: Show
		}
	}

	While (!InStr(WB.Document.location.href, "app=m")) {
		If (!DllCall("User32.dll\IsWindowVisible", "UInt", WinExist("ahk_id " hYouTube))) {
			return
		}

		Sleep, 250
	}

	If (!DllCall("User32.dll\IsWindowVisible", "UInt", WinExist("ahk_id " hAbout))) {
		Gui, App: -Disabled
	}

	Gui, YouTube: Hide
	SBSetText("Retrieving data...")

	While (WB.Document.getElementById("initial-data").innerHTML = "") {
		Sleep, 250
	}

	JSONid := WB.Document.getElementById("initial-data").innerHTML
	JSONid := RegExReplace(JSONid, "^<!--\s|\s-->$")
	JSONid := JavaEscapedToUnicode(JSONid)
	JSONigd := WB.Document.getElementById("initial-guide-data").innerHTML
	JSONigd := RegExReplace(JSONigd, "^<!--\s|\s-->$")
	JSONigd := JavaEscapedToUnicode(JSONigd)
	RegExMatch(JSONigd, """accountName"":.*?""text"":""(.*?)""}", AccountName)
	RegExMatch(JSONigd, """email"":.*?""text"":""(.*?)""}", AccountEmail)
	WB.document.documentElement.innerHTML := ""
	SBSetText("Idle", (AccountName1 && AccountEmail1 ? AccountName1 " <" AccountEmail1 ">" : ""))
	YT_Parse()
}

YT_Parse() {
	Global ; Assume-global mode

	SBSetText("Parsing subscription list...")
	Subscriptions := {}
	SubText := ""
	Pos := 0

	While (Pos := RegExMatch(JSONid, "{""channelListItemRenderer"":.*?""timestampMs"":""\d+""}}", Channel, Pos + 1)) {
		RegExMatch(Channel, """text"":""(.*?)""", ChannelName)
		RegExMatch(Channel, """canonicalBaseUrl"":""(.*?)""", ChannelURL)
		ChanName := ChannelName1
		ChanURL := "https://www.youtube.com" ChannelURL1
		ChanBaseURL := ChannelURL1
		Subscriptions[A_Index] := {"Name": ChanName, "URL": ChanURL, "BaseURL": ChanBaseURL}
	}

	For Index, Channel In Subscriptions {
		SBSetText("Building list (" Index " / " Subscriptions.MaxIndex() ")")
		SubText .= Channel.Name " - " Channel.URL (Index = Subscriptions.MaxIndex() ? "" : "`n")
	}

	GuiControl, App:, Edt, % SubText
	SBSetText("Idle")
	GuiControl, App: Enable, % hBtnStart
	WinActivate, ahk_id %hApp%
}

YT_Logout() {
	Global ; Assume-global mode

	SBSetText("Logging out...")
	WB.document.location.href := "https://m.youtube.com/logout?app=m&persist_app=1"

	While (!InStr(WB.Document.location.href, "noapp=1")) {
		Sleep, 250
	}

	WB.document.documentElement.innerHTML := ""
	SBSetText("Idle", "YouTube User")
}

YouTubeGuiSize(GuiHwnd, EventInfo, Width, Height) {
	Global ; Assume-global mode

	If (ErrorLevel = 1) { ; Window minimized
		return
	}

	GuiControl, Move, % hWB, % "w" Width " h" Height
}

AppGuiSize(GuiHwnd, EventInfo, Width, Height) {
	Global ; Assume-global mode

	If (ErrorLevel = 1) { ; Window minimized
		return
	}

	SB_SetParts(Width / 2)
	GuiControlGet, SB, Pos
	GuiControl, MoveDraw, % hBtnStart, % "y" (SBy - SBh) - 10
	GuiControlGet, BtnStart, Pos
	GuiControl, MoveDraw, % hEdt, % "w" Width - 20 " h" (BtnStarty - BtnStarth)
}

YouTubeGuiClose(GuiHwnd) {
	Global ; Assume-global mode

	Gui, App: -Disabled
	Gui, YouTube: Hide
	WB.document.documentElement.innerHTML := ""
	SBSetText("Idle", "YouTube User")
	GuiControl, Enable, % hBtnStart
}

YouTubeGuiEscape(GuiHwnd) {
	Global ; Assume-global mode

	Gui, App: -Disabled
	Gui, YouTube: Hide
	WB.document.documentElement.innerHTML := ""
	SBSetText("Idle", "YouTube User")
	GuiControl, Enable, % hBtnStart
}

AppTubeGuiClose(GuiHwnd) {
	Global ; Assume-global mode

	ExitApp
}

AppGuiEscape(GuiHwnd) {
	Global ; Assume-global mode

	ExitApp
}

AboutGuiClose(GuiHwnd) {
	Global ; Assume-global mode

	Gui, App: -Disabled
	Gui, About: Hide
}

AboutGuiEscape(GuiHwnd) {
	Global ; Assume-global mode

	Gui, App: -Disabled
	Gui, About: Hide
}

ControlHandler() {
	Global ; Assume-global mode

	If (A_GuiControl = "BtnStart") {
		YT_GetData()
	} Else If (A_GuiControl = "BtnAboutOK") {
		Gui, App: -Disabled
		Gui, About: Hide
	} Else If (A_GuiControl = "BtnAboutAHK") {
		Run, https://www.autohotkey.com/boards/memberlist.php?mode=viewprofile&u=56166
	} Else If (A_GuiControl = "BtnAboutGitHub") {
		Run, https://www.autohotkey.com/boards/memberlist.php?mode=viewprofile&u=56166
	} Else If (A_GuiControl = "SB") {
		If (A_GuiEvent = "RightClick") {
			If (A_EventInfo = "2") {
				StatusBarGetText, SBUserText, % A_EventInfo

				If (SBUserText = "") {
					return
				}

				Menu, SBUserMenu, Show, % A_GuiX, % A_GuiY
			}
		}
	}
}

MenuHandler(ItemName, ItemPos, MenuName) {
	Global ; Assume-global mode

	If (MenuName = "Tray") {
		If (ItemName = "E&xit") {
			ExitApp
		}
	} Else If (MenuName = "FileMenu") {
		If (InStr(ItemName, "E&xit")) {
			ExitApp
		}
	} Else If (MenuName = "HelpMenu") {
		If (InStr(ItemName, "&About")) {
			Gui, App: +Disabled
			Gui, About: Show
			GuiControl, Focus, % hBtnAboutOK
		}
	} Else If (MenuName = "SBUserMenu") {
		If (ItemName = "&Logout") {
			YT_Logout()
		} Else If (ItemName = "&Switch User") {
			YT_Logout()
			YT_GetData()
		}
	}
}

SBSetText(StrText*) {
	Global ; Assume-global mode

	PrevDefaultGui := A_DefaultGui
	Gui, App: Default

	For SBPart, SBText In StrText {
		SB_SetText(SBText, SBPart)
		SBText := RegExReplace(SBText, "`t")
		DllCall("User32.dll\SendMessage", "Ptr", hSB, "UInt", 1040 + (A_IsUnicode ? 1 : 0), "Ptr", SBPart - 1, "Ptr", &SBText)
	}

	Gui, %PrevDefaultGui%: Default
}

JavaEscapedToUnicode(s) {
	e := ""
	i := 1

	While (j := RegExMatch(s, "\\u[A-Fa-f0-9]{1,4}", m, i)) {
		e .= SubStr(s, i, j - i) Chr("0x" SubStr(m, 3))
		i := j + StrLen(m)
	}

	return e SubStr(s, i)
}

GdipCreateFromBase64(B64) {
	Global ; Assume-global mode

	VarSetCapacity(B64Len, 0)
	DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", StrLen(B64), "UInt", 0x01, "Ptr", 0, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
	VarSetCapacity(B64Dec, B64Len, 0) ; pbBinary size
	DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", StrLen(B64), "UInt", 0x01, "Ptr", &B64Dec, "UIntP", B64Len, "Ptr", 0, "Ptr", 0)
	pStream := DllCall("Shlwapi.dll\SHCreateMemStream", "Ptr", &B64Dec, "UInt", B64Len, "UPtr")
	VarSetCapacity(pBitmap, 0)
	DllCall("Gdiplus.dll\GdipCreateBitmapFromStreamICM", "Ptr", pStream, "PtrP", pBitmap)
	VarSetCapacity(hIcon, 0)
	DllCall("Gdiplus.dll\GdipCreateHICONFromBitmap", "Ptr", pBitmap, "PtrP", hIcon, "UInt", 0)
	ObjRelease(pStream)

	return hIcon
}
; ==============================================================================

; Classes ======================================================================
class WB_Events {
	DocumentComplete(WB) {
		WB.document.documentElement.style.setAttribute("overflow", "auto")
	}
}
; ==============================================================================