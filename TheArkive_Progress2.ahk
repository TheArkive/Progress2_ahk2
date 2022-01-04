; ahk v2
; ============================================================================================
; Progress bar window contains a MainText element (above the progress bar), a SubText element
; (below the progress bar), the progress bar itself, and an optional title bar.
;
; Here are the defaults:
;    Font                  = Verdana
;    Font Size (subText)   = 8
;    Font Size (mainText)  = subTextSize + 2
;   Main/SubText value    = blank unless specified
;    display coords        = primary monitor
;    default range         = 0-100
;
; ============================================================================================
; Create Progress Bar
;    obj := progress2.New(rangeStart := 0, rangeEnd := 100, sOptions := "")
;        > specify start and end range
;        > specify options to initialize certain values on creation (optional)
;
; Methods:
;
;    obj.Update(value := "", mainText := "", subText := "", range := "", title := "")
;        > value = # ... within the specified range.
;        > MainText / subText: update the text above / below the progress bar
;        > If you want to clear the mainText or subText pass a space, ie. " "
;        > To leave mainText / SubText unchanged, pass a zero-length string, ie. ""
;       >>> With this "range" parameter you must specify a range to change it.  A blank value will do nothing.
;       >>> With this "title" parameter you must specify a title to change it.  A blank value will do nothing.
;
;    obj.Close()
;        > closes the progress bar
;
;   obj.Range("#-#")
;       > Redefine the range on the fly.
;         Example:  obj.Range("0-50")
;       > Specify obj.Range("") to set the range to default (0-100).
;
; Properties:
;
;   obj.Title := "New Title"
;       > Change the title on the fly.  You can specify "" to clear the title.
;
;   obj.MainText := "new main text"
;       > Change Main Text (above the progress bar).
;
;   obj.SubText
;       > Change Sub Text (below the progress bar).
;
;   obj.Value
;       > Change the value of the progress bar.
;
;   obj.Range := "0-50"
;       > Change the range of the progress bar
;       > Pass "" to set range to default, ie. "0-100".
;
;   obj.RangeMin (read only)
;       > Get min range.
;
;   obj.RangeMax (read only)
;       > Get max range.
;
;   obj.hwnd
;       > Returns the hwnd of the GUI that contains the progress bar.
;
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
;    fontFace:font_str
;        > set the font for MainText / SubText
;
;    fontSize:###
;        > set font size for MainText / SubText (MainText is 2 pts larger than SubText)
;
;    mainText:text_str
;        > creates the Progress Bar with specified mainText (above the progress bar)
;
;    mainTextAlign:left/right/center
;        > specifies alignment of mainText element.
;
;    mainTextSize:#
;        > sets custom size for mainText element.
;
;    modal:Hwnd
;        > same as parent, but also disables the parent window while progressbar is active
;
;    parent:Hwnd
;        > defines the parent GUI window, and prevents a taskbar icon from appearing
;
;    start:#
;        > defines the starting numeric withing the specified range on creation
;
;    subText:text_str
;        > creates the Pogress Bar with specified subText (below the progress bar)
;
;    subTextAlign:left/right/center
;        > specifies alignment of subText element.
;
;    title:title_str
;        > Defines a title and ensures there is a title bar.  This allows normal moving of the
;          progress bar by click-dragging the title bar.  No title hides the title bar and
;          prevents the window from being moved by the mouse (by normal means).
;
;    w:###
;        > sets a specific pixel width for the progress bar
;
;    x:###  And  y:###  (specify both and separate by comma in options string)
;        > sets custom x/y coords to display the progress bar window.  Specify both or none.
;



; ============================================================================================
; Example
; ============================================================================================
; Global prog, clicks := 0

; g := Gui()
; g.OnEvent("close",close_gui)
; g.OnEvent("escape",close_gui)
; g.Add("Text","w600 h300","Test GUI")
; g.Add("Button",,"Test Progress - click 3 times slowly").OnEvent("click",click_btn)
; g.show("x200 y200")

; check_status() {
    ; Global prog
    ; If (prog.abort) {
        ; Msgbox "You clicked cancel."
        ; prog.abort := false ; This is only for the example, you would not normally do this.
    ; }                           ; Normally you use the Cancel button in a loop, and you keep checking for [ prog.abort = true ].
; }                               ; When you encounter [ prog.abort = true ] you decide how to handle it in your code.

; click_btn(p*) {
    ; Global prog, clicks
    
    ; If (clicks = 0) {
        ; options := "mainText:Test Main Text,subText:Test Sub Text,title:test title,"
        ; options .= "start:25,parent:" g.hwnd ",Cancel:1"
        ; prog := progress2(0,100,options)
        ; SetTimer check_status, 50
        ; clicks++
    ; } Else If (clicks = 1) {
        ; prog.Title := "Title change!!"
        ; prog.MainText := "Main Text change!!"
        ; prog.SubText := "Sub Text change!!"
        ; prog.Value := 50
        ; clicks++
    ; } Else if (clicks = 2) {
        ; Msgbox "Progress value: " prog.Value "`n" ; get property values
             ; . "Title: " prog.Title "`n"
             ; . "MainText: " prog.MainText "`n"
             ; . "SubText: " prog.SubText "`n"
             ; . "Range: " prog.Range "`n"
             ; . "Min range: " prog.RangeMin "`n"
             ; . "Max range: " prog.RangeMax
        
        ; prog.Update(75,"Main Text change again!","Sub Text change again!!",,"Title change again!!")
        ; Sleep 1000
        ; SetTimer check_status, 0
        ; prog.Close() ; Close with .Close() method.
        ; clicks := 0  ; Doing [ prog := "" ] won't work due to attached GUI and gui callback method.
    ; }                ; Of course prog can still be overwritten with another instance of itself.
; }

; close_gui(g) {
    ; ExitApp
; }
; ============================================================================================
; End Example
; ============================================================================================

class progress2 {
    rangeStart := 0, rangeEnd := 100        ; progress bar range
    fontFace := "Verdana", fontSize := 8    ; text Font and Size
    mainTextSize := 10                      ; set different text size of mainText (by default)
    width := 300                            ; GUI window width
    mainTextAlign := "left"                 ; alignment of mainText
    subTextAlign := "left"                  ; alignment of subText
    _title := ""                            ; GUI title
    _mainText := " ", _subText := " "       ; starting mainText / subText
    start := 0                              ; start value
    modal := false                          ; modal true/false
    parent := 0                             ; parent hwnd
    x := "", y := ""                        ; x,y of the GUI top-left corner
    AlwaysOnTop := 0
    cancel := false                         ; to use Cancel button or not
    abort := false                          ; has Cancel been clicked?
    hwnd := 0                               ; the GUI hwnd containing the progress bar
    
    __New(rangeStart := 0, rangeEnd := 100, sOptions := "") {
        this.rangeStart := rangeStart, this.rangeEnd := rangeEnd
        
        optArr := StrSplit(sOptions,Chr(44))
        Loop optArr.Length {
            valArr := StrSplit(optArr[A_Index],":")
            v := valArr[1]
            
            if RegExMatch(v,"i)^(title|mainText|subText)$")
                valArr[1] := "_" valArr[1]
            
            this.%valArr[1]% := valArr[2]
        }
        
        this.ShowProgress()
    }
    ShowProgress() {
        x := "", y := ""
        showTitle := this._title ? "" : " -Caption +0x40000" ; 0x40000 = thick border
        range := this.rangeStart "-" this.rangeEnd
        
        _styles := (this.AlwaysOnTop ? "AlwaysOnTop " : "") "-SysMenu " showTitle " +E0x02000000 +0x02000000"
        progress2_gui := Gui(_styles,this._title)
        this.hwnd := progress2_gui.hwnd
        
        progress2_gui.SetFont("s" this.mainTextSize,this.fontFace)
        align := this.mainTextAlign
        mT := progress2_gui.AddText("vMainText " align " w" this.width,this._mainText)
        
        progress2_gui.SetFont("s" this.fontSize)
        prog_ctl := progress2_gui.Add("Progress","vProgBar y+m xp w" this.width " Range" range,this.start)
        
        align := this.subTextAlign
        sT := progress2_gui.AddText("vSubText " align " w" this.width,this._subText)
        
        If (this.cancel) {
            btn := progress2_gui.Add("Button","vCancel w75 x" (this.width-75+progress2_gui.MarginX), "Cancel")
            progress2_gui.DefineProp("gui_events",{Call:(ctl,i) => ((ctl.name="Cancel")?(this.abort:=true):"")})
            btn.OnEvent("click", progress2_gui.gui_events) ; attach method to gui, not class
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
        
        progress2_gui.Show((!x && !y) ? "" : "x" x " y" y)
        
        this.x := x, this.y := y
        this.gui := progress2_gui
    }
    Update(value := "", mainText := "", subText := "", range := "", title := "") {
        If (value != "")
            this.gui["ProgBar"].Value := value
        If (mainText)
            this.gui["MainText"].Text := mainText
        If (subText)
            this.gui["SubText"].Text := subText
        If (range)
            this.gui["ProgBar"].Opt("Range" range)
        If (title)
            this.gui.Title := title
    }
    Title {
        set => this.gui.Title := value
        get => this.gui.Title
    }
    MainText {
        set => this.gui["MainText"].Text := value
        get => this.gui["MainText"].Text
    }
    SubText {
        set => this.gui["SubText"].Text := value
        get => this.gui["SubText"].Text
    }
    Value {
        set => this.gui["ProgBar"].Value := value
        get => this.gui["ProgBar"].Value
    }
    Range {
        set => this._SetRange(value)
        get => this.rangeStart "-" this.rangeEnd
    }
    RangeMin {
        get => this.rangeStart
    }
    RangeMax {
        get => this.rangeEnd
    }
    _SetRange(_in) {
        If !RegExMatch(_in,"^(\d+)\-(\d+)$",&m) && (_in != "")
            throw Error("Invalid range specified:`n`n" _in "`n`nProper format example:   '0-100'")
        this.rangeStart := (_in?m[1]:0)
        this.rangeEnd := (_in?m[2]:100)
        this.gui["ProgBar"].Opt("Range" this.rangeStart "-" this.rangeEnd)
    }
    Close() {
        If (IsObject(this.gui))
            this.gui.Destroy(), this.gui := "" ; completely delete this.gui to prevent destroying non-existent gui on __Delete()
        If (this.modal) {
            WinSetEnabled 1, "ahk_id " this.parent
            this.modal := this.parent := 0 ; prevent __Delete() from triggering this again
        }
    }
    __Delete() {
        this.Close()
    }
}