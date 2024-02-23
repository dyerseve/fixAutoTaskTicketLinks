; Example before url: https://ww14.autotask.net/Autotask/AutotaskExtend/ExecuteCommand.aspx?Code=OpenTicketDetail&TicketID=53203
; Example after url: https://ww14.autotask.net/Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&ids%5B0%5D=53203&ticketId=53203
#Requires AutoHotkey v2.0
#SingleInstance Force

OnClipboardChange CheckClipboard

CheckClipboard(DataType) {
    ;ToolTip "Clipboard data type: "  A_Clipboard
    clipboardContent := A_Clipboard
    match := ""
    ticketID := ""
    If (RegExMatch(clipboardContent, "https://ww\d+\.autotask\.net/Autotask/AutotaskExtend/ExecuteCommand\.aspx\?Code=OpenTicketDetail&TicketID=(\d+)", &match)) {
        ticketID := match[1]
        newURL := "https://ww14.autotask.net/Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&ids%5B0%5D=" ticketID "&ticketId=" ticketID
        ;ToolTip newURL
        A_Clipboard := newURL
    } else {
        ; If no match found, do nothing or handle it as needed
    }
}
