; v0.6 Moved url decodes to behind a hotkey, to avoid accidental usage when not intended. URLDecode Clipboard Hotkey = CTRL+?
; v0.5 Added single encode for barracuda as seen in service tickets
; v0.4 Decode double "protected" links from barracuda and microsoft.
; v0.3 Added Ticket Number search
; v0.2 Added start and stop string as copying the sections below to make copies was causing it to find the substring and replace my clipboard, doh!
; ExecuteCommand only opens a popup, so can't use that, some tasks are tickets, some are tasks in Projects. A task in Project is less used and tasks in general are less used so I gave up trying to find any way to discern from the ExecuteCommand which type of task it is.

; Before: Share ticket url: https://ww14.autotask.net/Autotask/AutotaskExtend/ExecuteCommand.aspx?Code=OpenTicketDetail&TicketID=53203
; After: Share ticket url: https://ww14.autotask.net/Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&ids%5B0%5D=53203&ticketId=53203

; Before: Project Task: https://ww14.autotask.net/Autotask/AutotaskExtend/ExecuteCommand.aspx?Code=OpenTaskDetail&taskid=45088
; After: Project Task: https://ww14.autotask.net/Mvc/Projects/TaskDetail.mvc?taskID=45088

; Before: Convert ticket number to link: T20240223.0015
; After: Convert ticket number to link: https://ww14.autotask.net/Mvc/ServiceDesk/TicketGridSearch.mvc/SearchByTicketNumber?TicketNumber=T20240223.0015

; This one only converts on hotkey press CTRL+?(CTRL+SHIFT+/)
; This is purposely meant to be difficult to do because these wraps are part of the protection of these links, they are not intended to be accessed directly.

; Before: Office Protect link: https://nam04.safelinks.protection.outlook.com/?url=https%3A%2F%2Flinkprotect.cudasvc.com%2Furl%3Fa%3Dhttps%253a%252f%252fsmotwildcats.org%26c%3DE%2C1%2COeb6MxVttUeyj2Str2IvhoQRRnnfPxWWkZ0GkZD8vdXZ91gRceQRQjTYH5OSZvXrsHloTISrHzYsWtm351zg4xgB2ax8NFaz5-y7aic_tevKwQ4fngrlB1JX2JCg%26typo%3D1&data=05%7C02%7Cpellis%40cmitsolutions.com%7C9f8f0066490c41fc7b2608dc3c5a566e%7C7bf5c1815bc4497e9db53845ce848a0d%7C0%7C0%7C638451605914421018%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&sdata=hiH3Queo0Rx8r%2BT7UQltM5CtB%2BXTQvs%2B0ytpK%2BdLQxM%3D&reserved=0
; After: Office Protect link: https://smotwildcats.org/

; Sometimes our ticketing system will double wrap the weblinks a customer sends up. Often, I just need to know what it is, so decoding it in clipboard seems ideal.

; This one only converts on hotkey press
; Before: cudaprotect links: https://linkprotect.cudasvc.com/url?a=https%3a%2f%2fsecure.cpacharge.com%2fpages%2friesenbeckfinancialgroupinc%2fpayments&c=E,1,80Yo5E_P0a-IBcQRbqnhA_bLuZ-5k6jRjR8Sj4yHRNHSkSb8UAQQBJD93BZAIa5B6ylopCxxbyquzKUFogjXDJOMjNIfvrIunA3W8ykAgWIxO3Y,&typo=1
; After: cudaprotect links: https://secure.cpacharge.com/pages/riesenbeckfinancialgroupinc/payments


#Requires AutoHotkey v2.0
#SingleInstance Force

OnClipboardChange CheckClipboard
^?::
{
    try {
        clipboardData := A_Clipboard
    } 
    catch {
        ; Ignore the error and continue
    }
    match := ""
    ticketID := ""
    ; Office Protect link start
    If (RegExMatch(clipboardData, "^https:\/\/nam04\.safelinks\.protection\.outlook\.com\/\?url=https%3A%2F%2Flinkprotect.cudasvc.com%2Furl%3Fa%3D(http.*?)%26c%3DE%2C1%2C", &match)) {
        encurl := match[1]
        encurl := UrlUnescape(encurl)
        encurl := UrlUnescape(encurl)
        newURL := encurl
        ;ToolTip newURL
        A_Clipboard := newURL
    } else {
        ; If no match found, do nothing or handle it as needed
    }
    ; Office Protect link end
    ; cudaprotect links start
    If (RegExMatch(clipboardData, "^https:\/\/linkprotect\.cudasvc\.com\/url\?a=(http.*?)\&c=E,1,", &match)) {
        encurl := match[1]
        encurl := UrlUnescape(encurl)
        newURL := encurl
        ;ToolTip newURL
        A_Clipboard := newURL
    } else {
        ; If no match found, do nothing or handle it as needed
    }
    ; cudaprotect links end
}

CheckClipboard(DataType) {
    try {
        clipboardData := A_Clipboard
    } 
    catch {
        ; Ignore the error and continue
    }
    match := ""
    ticketID := ""
    ; Share ticket url start
    If (RegExMatch(clipboardData, "^https://ww\d+\.autotask\.net/Autotask/AutotaskExtend/ExecuteCommand\.aspx\?Code=OpenTicketDetail&TicketID=(\d+)$", &match)) {
        ticketID := match[1]
        newURL := "https://ww14.autotask.net/Mvc/ServiceDesk/TicketDetail.mvc?workspace=False&ids%5B0%5D=" ticketID "&ticketId=" ticketID
        ;ToolTip newURL
        A_Clipboard := newURL
    } else {
        ; If no match found, do nothing or handle it as needed
    }
    ; Share ticket url end
    ; Convert ticket number to link start
    If (RegExMatch(clipboardData, "^(T20\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\.\d{4})$", &match)) {
        ticketID := match[1]
        newURL := "https://ww14.autotask.net/Mvc/ServiceDesk/TicketGridSearch.mvc/SearchByTicketNumber?TicketNumber=" ticketID
        ;ToolTip newURL
        A_Clipboard := newURL
    } else {
        ; If no match found, do nothing or handle it as needed
    }
    ; Convert ticket number to link end
    ; Project Task start
    If (RegExMatch(clipboardData, "^https://ww\d+\.autotask\.net/Autotask/AutotaskExtend/ExecuteCommand\.aspx\?Code=OpenTaskDetail&taskid=(\d+)$", &match)) {
        taskID := match[1]
        newURL := "https://ww14.autotask.net/Mvc/Projects/TaskDetail.mvc?taskID=" taskID
        ;ToolTip newURL
        A_Clipboard := newURL
    } else {
        ; If no match found, do nothing or handle it as needed
    }
    ; Project Task end
    

}

UrlEscape(Url, Flags := 0x000C1000) {
   Static CC := 4096
   VarSetStrCapacity(&Esc, CC)
   Return !DllCall("Shlwapi.dll\UrlEscape", "Str", Url, "Str", Esc, "UIntP", &CC, "UInt", Flags, "UInt") ? Esc : ""
}

UrlUnescape(Url, Flags := 0x00140000) {
   Return !DllCall("Shlwapi.dll\UrlUnescape", "Ptr", StrPtr(Url), "Ptr", 0, "UInt", 0, "UInt", Flags, "UInt") ? Url : ""
}
