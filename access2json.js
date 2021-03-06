// Generated by CoffeeScript 1.6.3
/*
    Сначала подключаем базовые классы
*/


(function() {
  var A2JsonReader, Access2json, Fs, a2js, dbReaderFile, programFile, _ref,
    __hasProp = {}.hasOwnProperty,
    __extends = function(child, parent) { for (var key in parent) { if (__hasProp.call(parent, key)) child[key] = parent[key]; } function ctor() { this.constructor = child; } ctor.prototype = parent.prototype; child.prototype = new ctor(); child.__super__ = parent.prototype; return child; };

  Fs = new ActiveXObject("Scripting.FileSystemObject");

  this.Folder = Fs.GetParentFolderName(Fs.GetFile(WScript.ScriptFullName)) + "\\";

  programFile = this.Folder + "program.js";

  if (!Fs.FileExists(programFile)) {
    WScript.Echo("program.js file not found");
    WScript.Quit(1);
  }

  dbReaderFile = this.Folder + "accessDbReader.js";

  if (!Fs.FileExists(dbReaderFile)) {
    WScript.Echo("AccessDbReader file not found");
    WScript.Quit(1);
  }

  eval(Fs.OpenTextFile(programFile, 1).ReadAll());

  eval(Fs.OpenTextFile(dbReaderFile, 1).ReadAll());

  /*
      ОК. Базовые классы теперь доступны.
  
      A2JsonReader складывает данные в массив - для обьектов первого уровня конфига
      Используйте pretty опцию конструктора для форматирования с использованием сдвига(для удобства отладки)
  */


  A2JsonReader = (function(_super) {
    __extends(A2JsonReader, _super);

    A2JsonReader.OutFile = null;

    A2JsonReader.pretty = true;

    function A2JsonReader(out, pretty) {
      this.OutFile = out;
      this.pretty = pretty;
    }

    A2JsonReader.prototype.log = function(str) {
      var t;
      t = void 0;
      if (typeof str === "undefined") {
        t = "undefined";
      } else if (str === null) {
        t = "null";
      } else {
        t = str;
      }
      if (this.OutFile) {
        return this.OutFile.Write(t);
      } else {
        return WScript.StdOut.Write(t);
      }
    };

    A2JsonReader.prototype.process = function(state, level, data) {
      var addIndent, self;
      if (this.pretty) {
        self = this;
        addIndent = function(nl) {
          var i;
          i = nl || level;
          while (i > 0) {
            self.log("\t");
            i--;
          }
        };
      } else {
        addIndent = function() {
          return {};
        };
      }
      switch (state) {
        case this.ADBR_ARRAY_START:
          addIndent();
          this.log("[");
          if (this.pretty) {
            return this.log("\n");
          }
          break;
        case this.ADBR_ARRAY_END:
          if (this.pretty) {
            this.log("\n");
          }
          addIndent();
          return this.log("]");
        case this.ADBR_OBJ_START:
          addIndent();
          this.log("{");
          if (this.pretty) {
            return this.log("\n");
          }
          break;
        case this.ADBR_OBJ_END:
          if (this.pretty) {
            this.log("\n");
          }
          addIndent();
          return this.log("}");
        case this.ADBR_NEXT_OBJ:
        case this.ADBR_NEXT_PROP:
          this.log(",");
          if (this.pretty) {
            return this.log("\n");
          }
          break;
        case this.ADBR_PROP_NAME:
          addIndent(level + 1);
          return this.log('"' + data + '":');
        case this.ADBR_PROP_VALUE:
          return this.log(JSON.stringify(data));
        case this.ADBR_BEFORE_CHILD:
          if (this.pretty) {
            return this.log("\n");
          }
      }
    };

    return A2JsonReader;

  })(AccessDbReader);

  /*
      Используем общую часть (инициализацию описанную в классе Program )
      Готовим новый файл (затираем существующий) в который будем складывать результат
      Используем A2JsonReader для представления данных
  */


  Access2json = (function(_super) {
    __extends(Access2json, _super);

    function Access2json() {
      _ref = Access2json.__super__.constructor.apply(this, arguments);
      return _ref;
    }

    Access2json.prototype.prepareOutFile = function() {
      var f;
      f = this.Folder + "a2json_result.json";
      return this.OutFile = Fs.CreateTextFile(f, 8, true);
    };

    Access2json.prototype.run = function() {
      var e, r;
      try {
        r = new A2JsonReader(this.OutFile, true);
        return r.read(this.dbPath, this.Config);
      } catch (_error) {
        e = _error;
        WScript.Echo(e.name + "\n\n" + e.description);
        return WScript.Quit(2);
      }
    };

    return Access2json;

  })(Program);

  a2js = new Access2json();

  a2js.run();

}).call(this);

/*
//@ sourceMappingURL=access2json.map
*/
