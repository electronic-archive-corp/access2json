#force to use cscript and if arch is 64 bit - use correct cscript to start

((ws) ->
    callAndWait = (execStr) ->
        oExec = WshShell.Exec(execStr)
        TextStream = oExec.StdOut
        while oExec.Status == 0
            while(!TextStream.AtEndOfStream)
                ws.Echo(TextStream.ReadLine())
                #WScript.Sleep(10);
        return
        #ws.Echo TextStream.ReadLine()  until TextStream.AtEndOfStream  while oExec.Status is 0

    WshShell = WScript.CreateObject("WScript.Shell")
    arch = WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    cmd = "cscript.exe //nologo \"" + ws.scriptFullName + "\""
    correctArch = true
    if arch.indexOf("86") < 0
        sr = WshShell.ExpandEnvironmentStrings("%SystemRoot%")
        cmd = sr + "\\SysWOW64\\" + cmd
        correctArch = false

    args = ws.arguments
    i = 0
    len = args.length

    while i < len
        arg = args(i)
        cmd += " " + ((if ~arg.indexOf(" ") then "\"" + arg + "\"" else arg))
        i++
        #cmd += ' 2>&1';

    isCScriptUsed = true || ws.fullName.slice(-12).toLowerCase() isnt "\\cscript.exe"
    if !correctArch or !isCScriptUsed
        callAndWait cmd
        ws.quit()
) WScript