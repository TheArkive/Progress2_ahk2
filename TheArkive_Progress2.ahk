; ahk v2
; ============================================================================================
; Example
; ============================================================================================
; class app {
    ; Static prog := ""
; }

; g := Gui()
; g.OnEvent("close",close_gui)
; g.Add("Text","w600 h300","Test GUI")
; g.Add("Button",,"Test Progress - click 2 times slowly").OnEvent("click",click_btn)
; g.show("x200 y200")

; click_btn(p*) {
	; If (!IsObject(app.prog)) {
		; options := "mainText:Test Main Text pqg,subText:Test Sub Text pqg,title:test title,"
		; options .= "start:25,parent:" g.hwnd ",Cancel:1"
		; app.prog := progress2(0,100,options)
	; } Else {
        ; msgbox "Abort status: " app.prog.abort
        
		; app.prog.Update(50," ","Test 3")
		; Sleep 2000
        ; app.prog.Close()
        ; app.prog := "" ; needed for this example to work properly
	; }
; }

; close_gui(g) {
	; ExitApp
; }
; ============================================================================================
; End Example
; ============================================================================================


; ============================================================================================
; Progress bar window contains a MainText element (above the progress bar), a SubText element
; (below the progress bar), the progress bar itself, and an optional title bar.
;
; Here are the defaults:
;	Font                  = Verdana
;	Font Size (subText)   = 8
;	Font Size (mainText)  = subTextSize + 2
;   Main/SubText value    = blank unless specified
;	display coords        = primary monitor
;	default range         = 0-100
;
; ============================================================================================
; Create Progress Bar
;	obj := progress2.New(rangeStart := 0, rangeEnd := 100, sOptions := "")
;		> specify start and end range
;		> specify options to initialize certain values on creation (optional)
;
; Methods:
;	obj.Update(value := "", mainText := "", subText := "", range := "", title := "")
;		> value = # ... within the specified range.
;		> MainText / subText: update the text above / below the progress bar
;		> If you want to clear the mainText or subText pass a space, ie. " "
;		> To leave mainText / SubText unchanged, pass a zero-length string, ie. ""
;       >>> With this "range" parameter you must specify a range to change it, unlike "obj.Range()" below.
;       >>> With this "title" parameter you must specify a title to change it, unlike "obj.Title()" below.
;
;	obj.Close()
;		> closes the progress bar
;
;   obj.Range("#-#")
;       > Redefine the range on the fly.
;         Example:  obj.Range("0-50")
;       > Specify obj.Range("") to set the range to default (0-100).
;
;   obj.Title(newTitle)
;       > Change the title on the fly.  You can specify "" to clear the title.
; ============================================================================================
; sOptions on create.  Comma separated string including zero or more of the below options.
; MainText / SubText are the text above / below the progress bar.
;
;   AlwaysOnTop:1
;       > set the progress bar to be always on top of other windows.
;
;   Cancel:1
;       > Add a cancel button.  Clicking the button will set obj.abort := true
;         It is up to the coder to decide how to check obj.abort, and how to
;         respond when obj.abort = true
;
;	fontFace:font_str
;		> set the font for MainText / SubText
;
;	fontSize:###
;		> set font size for MainText / SubText (MainText is 2 pts larger than SubText)
;
;	mainText:text_str
;		> creates the Progress Bar with specified mainText (above the progress bar)
;
;	mainTextAlign:left/right/center
;		> specifies alignment of mainText element.
;
;	mainTextSize:#
;		> sets custom size for mainText element.
;
;	modal:Hwnd
;		> same as parent, but also disables the parent window while progressbar is active
;
;	parent:Hwnd
;		> defines the parent GUI window, and prevents a taskbar icon from appearing
;
;	start:#
;		> defines the starting numeric withing the specified range on creation
;
;	subText:text_str
;		> creates the Pogress Bar with specified subText (below the progress bar)
;
;	subTextAlign:left/right/center
;		> specifies alignment of subText element.
;
;	title:title_str
;		> Defines a title and ensures there is a title bar.  This allows normal moving of the
;		  progress bar by click-dragging the title bar.  No title hides the title bar and
;		  prevents the window from being moved by the mouse (by normal means).
;
;	width:###
;		> sets a specific pixel width for the progress bar
;
;	x:###  And  y:###  (specify both and separate by comma in options string)
;		> sets custom x/y coords to display the progress bar window.  Specify both or none.
;
; ============================================================================================

class progress2 {
    rangeStart := 0, rangeEnd := 100
    fontFace := "Verdana", fontSize := 8, mainTextSize := 10
    mainTextAlign := "left", subTextAlign := "left"
    mainText := " ", subText := " ", title := "", start := 0, modal := false, parent := 0, width := 300
    x := "", y := "", mainTextHwnd := 0, subTextHwnd := 0, AlwaysOnTop := 0, cancel := false
    abort := false
    
	__New(rangeStart := 0, rangeEnd := 100, sOptions := "") {
		this.rangeStart := rangeStart, this.rangeEnd := rangeEnd
		
		optArr := StrSplit(sOptions,Chr(44))
        Loop optArr.Length {
            valArr := StrSplit(optArr[A_Index],":")
            this.%valArr[1]% := valArr[2]
        }
        
		this.ShowProgress()
	}
	ShowProgress() {
        x := "", y := ""
		showTitle := this.title ? "" : " -Caption +0x40000" ; 0x40000 = thick border
		range := this.rangeStart "-" this.rangeEnd
        
		progress2_gui := Gui((this.AlwaysOnTop ? "AlwaysOnTop " : "") "-SysMenu " showTitle,this.title)
		
		progress2_gui.SetFont("s" this.mainTextSize,this.fontFace)
		align := this.mainTextAlign
		mT := progress2_gui.Add("Text","vMainText r1 " align " w" this.width,this.mainText)
		this.mainTextHwnd := mT.hwnd
		
        progress2_gui.SetFont("s" this.fontSize)
		prog_ctl := progress2_gui.Add("Progress","vProgBar y+m xp w" this.width " Range" range,this.start)
		this.progHwnd := prog_ctl.hwnd
		
		align := this.subTextAlign
		sT := progress2_gui.Add("Text","vSubText r1 " align " w" this.width,this.subText)
		this.subTextHwnd := sT.hwnd
		
        If (this.cancel) {
            btn := progress2_gui.Add("Button","vCancel w75 x" (this.width-75+progress2_gui.MarginX), "Cancel")
            btn.OnEvent("click", ObjBindMethod(this,"gui_events"))
        }
        
		If (this.parent) {
			WinGetPos &pX, &pY, &pW, &pH, "ahk_id " this.parent
			Cx := pX + (pW/2), Cy := pY + (pH/2)
			progress2_gui.Opt("+Owner" this.parent)
			
			If (this.modal)
				WinSetEnabled 0, "ahk_id " this.parent
		}
		progress2_gui.Show(" NA NoActivate Hide") ; coords ??
        progress2_gui.GetPos(,,&w,&h)
        
        If (this.x = "" Or this.y = "") And this.parent
            x := Cx - (w/2), y := Cy - (h/2)
        
        progress2_gui.Show("x" x " y" y)
        
        this.x := x, this.y := y
		this.guiHwnd := progress2_gui.hwnd
		this.gui := progress2_gui
	}
    gui_events(ctl, info) {
        If (ctl.name = "Cancel")
            this.abort := true
    }
	Update(value := "", mainText := "", subText := "", range := "", title := "") {
		If (value != "")
			this.gui["ProgBar"].Value := value
		If (this.mainTextHwnd And mainText)
			this.gui["MainText"].Text := mainText
		If (this.subTextHwnd And subText)
			this.gui["SubText"].Text := subText
        If (range)
            this.gui["ProgBar"].Opt("Range" range)
        If (title)
            this.gui["ProgBar"].Title := title
	}
    Range(range := "0-100") {
        this.gui["ProgBar"].Opt("Range" range)
    }
    SetTitle(newTitle := "") {
        this.gui["ProgBar"].Title := newTitle
    }
	Close() {
		If (IsObject(this.gui))
			this.gui.Destroy(), this.gui := ""
		If (this.modal)
			WinSetEnabled 1, "ahk_id " this.parent
	}
}