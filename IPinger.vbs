lol=msgbox("This will display your IP Address to everyone looking at your computer.",64,"IPinger")
ReturnValue = MsgBox("Click yes to exit." + vbCrlf + "Click no to continue.", 36, "Exit?")
If ReturnValue = 7 Then
dim NIC1, Nic, StrIP, CompName

Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1

    if Nic.IPEnabled then
        StrIP = Nic.IPAddress(0)

        Set WshNetwork = WScript.CreateObject("WScript.Network")
        CompName= WshNetwork.Computername

        MsgBox "IP Address:  "&StrIP & vbNewLine _
            & "Computer Name:  "&CompName,4160,"IP Address and Computer Name"

        wscript.quit
    End if
Next
If ReturnValue = 6 Then
    MsgBox "Script is now exiting"
    WScript.Quit
End if
End If
