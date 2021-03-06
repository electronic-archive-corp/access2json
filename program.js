// Generated by CoffeeScript 1.6.3
(function() {
  var self;

  self = this;

  /*
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
  */


  this.Program = (function() {
    Program.prototype.Folder = null;

    Program.prototype.OutFile = null;

    Program.prototype.Config = null;

    Program.prototype.dbPath = null;

    Program.prototype.confPath = null;

    function Program() {
      var CONFIG, Fs, JSON, correctCScript, jsonFile, objArgs;
      Fs = new ActiveXObject("Scripting.FileSystemObject");
      this.Folder = Fs.GetParentFolderName(Fs.GetFile(WScript.ScriptFullName)) + "\\";
      correctCScript = this.Folder + "forceCorrectCScriptUse.js";
      if (Fs.FileExists(correctCScript)) {
        eval(Fs.OpenTextFile(correctCScript, 1).ReadAll());
      }
      jsonFile = this.Folder + "json2.js";
      if (!Fs.FileExists(jsonFile)) {
        WScript.Echo("json2.js file not found");
        this.Quit(1);
      }
      JSON = {};
      eval(Fs.OpenTextFile(jsonFile, 1).ReadAll());
      if (typeof JSON === "undefined" || typeof JSON.stringify === "undefined") {
        WScript.Echo("JSON object is not valid");
        this.Quit(1);
      }
      self.JSON = JSON;
      this.dbPath = this.Folder + "mdb.mdb";
      this.confPath = this.Folder + "config.json";
      objArgs = WScript.Arguments;
      if (objArgs.length >= 1) {
        if (objArgs[0] !== "null") {
          this.dbPath = objArgs[0];
        }
      }
      if (objArgs.length > 1) {
        if (objArgs[1] !== "null") {
          this.confPath = objArgs[1];
        }
      }
      this.prepareOutFile();
      if (!Fs.FileExists(this.dbPath)) {
        WScript.Echo("Database file " + this.dbPath + " not found");
        this.Quit(1);
      }
      if (!Fs.FileExists(this.confPath)) {
        WScript.Echo("Config file " + this.confPath + " not found");
        this.Quit(1);
      }
      CONFIG = {};
      eval("CONFIG=" + Fs.OpenTextFile(this.confPath, 1).ReadAll());
      this.Config = CONFIG;
      if (typeof this.Config === "undefined" || (this.Config == null)) {
        WScript.Echo("Config is not valid");
        this.Quit(1);
      }
    }

    Program.prototype.prepareOutFile = function() {};

    Program.prototype.run = function() {};

    Program.prototype.Quit = function(code) {
      if (this.OutFile) {
        this.OutFile.Close();
      }
      return WScript.Quit(code);
    };

    return Program;

  })();

}).call(this);

/*
//@ sourceMappingURL=program.map
*/
