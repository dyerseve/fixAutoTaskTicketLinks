; v0.2 Added start and stop string as copying the sections below to make copies was causing it to find the substring and replace my clipboard, doh!
; ExecuteCommand only opens a popup, so can't use that, some tasks are tickets, some are tasks in Projects. A task in Project is less used and tasks in general are less used so I gave up trying to find any way to discern from the ExecuteCommand which type of task it is.
 
; Example ticket before url: https://ww14.autotask.net/Autotask/AutotaskExtend/ExecuteCommand.aspx?Code=OpenTicketDetail&TicketID=53203
; Example ticket after url: https://ww14.autotask.net/Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&ids%5B0%5D=53203&ticketId=53203

;Project Task: https://ww14.autotask.net/Autotask/AutotaskExtend/ExecuteCommand.aspx?Code=OpenTaskDetail&taskid=45088
;Project Task: https://ww14.autotask.net/Mvc/Projects/TaskDetail.mvc?taskID=45088

;Project Task: https://ww14.autotask.net/Autotask/AutotaskExtend/ExecuteCommand.aspx?Code=OpenTaskDetail&taskid=48429
;Project Task: https://ww14.autotask.net/Mvc/Projects/TaskDetail.mvc?taskID=48429

;Project: https://ww14.autotask.net/Autotask/AutotaskExtend/ExecuteCommand.aspx?Code=OpenTaskDetail&taskid=48934
;Project: https://ww14.autotask.net/Mvc/ServiceDesk/TicketDetail.mvc?ticketId=48934

#Requires AutoHotkey v2.0
#SingleInstance Force

OnClipboardChange CheckClipboard

CheckClipboard(DataType) {
    ;ToolTip "Clipboard data type: "  A_Clipboard
    clipboardContent := A_Clipboard
    match := ""
    ticketID := ""
    If (RegExMatch(clipboardContent, "^https://ww\d+\.autotask\.net/Autotask/AutotaskExtend/ExecuteCommand\.aspx\?Code=OpenTicketDetail&TicketID=(\d+)$", &match)) {
        ticketID := match[1]
        newURL := "https://ww14.autotask.net/Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&ids%5B0%5D=" ticketID "&ticketId=" ticketID
        ;ToolTip newURL
        A_Clipboard := newURL
    } else {
        ; If no match found, do nothing or handle it as needed
    }
}
