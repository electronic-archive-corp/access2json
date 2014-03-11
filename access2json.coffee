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

    A2JsonReader складывает данные в массив - для обьектов первого уровня конфига
    Используйте pretty опцию конструктора для форматирования с использованием сдвига(для удобства отладки)
###
class A2JsonReader extends AccessDbReader
    @OutFile: null
    @pretty: true

    constructor: (out, pretty)->
        @OutFile = out
        @pretty = pretty


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
        if @pretty
            self = this
            addIndent = (nl)->
                i = nl || level
                while i > 0
                    self.log("\t")
                    i--
                return
        else
            addIndent = ()->{}

        switch state
            when @ADBR_ARRAY_START
                addIndent()
                @log "["
                if @pretty
                    @log "\n"
            when @ADBR_ARRAY_END
                if @pretty
                    @log "\n"
                addIndent()
                @log "]"
            when @ADBR_OBJ_START
                addIndent()
                @log "{"
                if @pretty
                    @log "\n"
            when @ADBR_OBJ_END
                if @pretty
                    @log "\n"
                addIndent()
                @log "}"
            when @ADBR_NEXT_OBJ, @ADBR_NEXT_PROP
                @log ","
                if @pretty
                    @log("\n")
            when @ADBR_PROP_NAME
                addIndent(level + 1)
                @log '"' + data + '":'
            when @ADBR_PROP_VALUE
                @log JSON.stringify data
            when @ADBR_BEFORE_CHILD
                if @pretty
                    @log("\n")
            #when @ADBR_AFTER_CHILD


###
    Используем общую часть (инициализацию описанную в классе Program )
    Готовим новый файл (затираем существующий) в который будем складывать результат
    Используем A2JsonReader для представления данных
###
class Access2json extends Program
    prepareOutFile: ()->
        f = @Folder + "a2json_result.json"
        @OutFile = Fs.CreateTextFile(f, 8, true)

    run: ()->
        try
            r = new A2JsonReader(@OutFile, true)
            r.read(@dbPath, @Config)
        catch e
            WScript.Echo(e.name + "\n\n" + e.description)
            WScript.Quit(2)


# запуск программы
a2js = new Access2json()
a2js.run();
