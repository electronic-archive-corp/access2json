###
    Сначала подключаем базовые классы
###
Fs = new ActiveXObject("Scripting.FileSystemObject")
@Folder = Fs.GetParentFolderName(Fs.GetFile(WScript.ScriptFullName)) + "\\"
programFile = @Folder + "program.js"
unless Fs.FileExists(programFile)
    WScript.Echo "program.js file not found"
    WScript.Quit 1

dbReaderFile = @Folder + "accessDbReader.js"
unless Fs.FileExists(dbReaderFile)
    WScript.Echo "AccessDbReader file not found"
    WScript.Quit 1

eval Fs.OpenTextFile(programFile, 1).ReadAll()
eval Fs.OpenTextFile(dbReaderFile, 1).ReadAll()


###
    ОК. Базовые классы теперь доступны.

    A2BsonReader складывает данные в одну строчку на обьект для обьектов первого уровня конфига
    Разделение между обьектами - новая строка
###
class A2BsonReader extends AccessDbReader
    @OutFile: null

    constructor: (out, p)->
        @OutFile = out


    log: (str) ->
        t = undefined
        if typeof (str) is "undefined"
            t = "undefined"
        else if str is null
            t = "null"
        else
            t = str
        if @OutFile
            @OutFile.Write t
        else
            WScript.StdOut.Write t


    process: (state, level, data)->
        switch state
            when @ADBR_ARRAY_START
                if(level != 0)
                    @log "["
            when @ADBR_ARRAY_END
                if(level != 0)
                    @log "]"

            when @ADBR_OBJ_START
                @log "{"
            when @ADBR_OBJ_END
                @log "}"
            when @ADBR_NEXT_OBJ
                @log ","
                if(level == 1)
                    @log "\n"
            when @ADBR_NEXT_PROP
                @log ","
            when @ADBR_PROP_NAME
                @log '"' + data + '":'
            when @ADBR_PROP_VALUE
                @log JSON.stringify data


###
    Используем общую часть (инициализацию описанную в классе Program )
    Готовим новый файл (затираем существующий) в который будем складывать результат
    Используем A2BsonReader для представления данных
###
class Access2bson extends Program
    prepareOutFile: ()->
        f = @Folder + "a2bson_result.json"
        @OutFile = Fs.CreateTextFile(f, 8, true)


    run: ()->
        try
            r = new A2BsonReader(@OutFile)
            r.read(@dbPath, @Config)
        catch e
            WScript.Echo(e.name + "\n\n" + e.description)
            WScript.Quit(2)


# запуск программы
a2js = new Access2bson()
a2js.run();
