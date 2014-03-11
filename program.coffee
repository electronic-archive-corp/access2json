# используется здесь для того чтобы глобальному обьекту прицепить JSON
self = this;

###
    Класс для загрузки всех необходимых для работы программы обьектов, их первоначальной проверки и инициализации
        - JSON
        - использование правильной версии cscript (версия 64 бит не совместима с драйвером аccess)
        - проверка доступности базы данных
        - загрузка конфига

    Все действия происходят в конструкторе класса.
    Тот кто планирует использовать этот класс должен определить 2 метода
        prepareOutFile:
            Создать файл в который будет складывать результаты своей работы
        run:
            Выбрать себе ридер базы данных, проинициализировать, обработать ошибки
###
class @Program
    Folder: null
    OutFile: null
    Config: null
    dbPath: null
    confPath: null

    constructor: ()->
        Fs = new ActiveXObject("Scripting.FileSystemObject")
        @Folder = Fs.GetParentFolderName(Fs.GetFile(WScript.ScriptFullName)) + "\\"

        correctCScript = @Folder + "forceCorrectCScriptUse.js"
        if Fs.FileExists correctCScript
            eval Fs.OpenTextFile(correctCScript, 1).ReadAll()

        jsonFile = @Folder + "json2.js"
        unless Fs.FileExists(jsonFile)
            WScript.Echo "json2.js file not found"
            @Quit 1
        JSON = {}
        eval Fs.OpenTextFile(jsonFile, 1).ReadAll()
        if typeof (JSON) is "undefined" or typeof (JSON.stringify) is "undefined"
            WScript.Echo "JSON object is not valid"
            @Quit 1
        self.JSON = JSON
        @dbPath = @Folder + "mdb.mdb"
        @confPath = @Folder + "config.json"
        objArgs = WScript.Arguments
        if objArgs.length >= 1
            unless objArgs[0] is "null"
                @dbPath = objArgs[0]
                #log "using dbPath: " + dbPath

        if objArgs.length > 1
            unless objArgs[1] is "null"
                @confPath = objArgs[1]

        @prepareOutFile()

        unless Fs.FileExists(@dbPath)
            WScript.Echo "Database file " + @dbPath + " not found"
            @Quit 1
        unless Fs.FileExists(@confPath)
            WScript.Echo "Config file " + @confPath + " not found"
            @Quit 1

        CONFIG = {}
        eval "CONFIG=" + Fs.OpenTextFile(@confPath, 1).ReadAll()
        @Config = CONFIG
        if typeof (@Config) is "undefined" or not @Config?
            WScript.Echo "Config is not valid"
            @Quit 1


    prepareOutFile: ()->
        #override
    run: () ->
        #override


    Quit: (code) ->
        if @OutFile
            @OutFile.Close()
        WScript.Quit(code)