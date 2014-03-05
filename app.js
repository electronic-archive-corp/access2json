var WshShell = WScript.CreateObject ("WScript.Shell");
var Fs = new ActiveXObject("Scripting.FileSystemObject");

function alert(str){
	WScript.Echo(str);
}

// this one is directory of process
//alert (WshShell.CurrentDirectory);

//find current directory of this file
var objFile = Fs.GetFile(WScript.ScriptFullName);
var Folder = Fs.GetParentFolderName(objFile) + "\\";

var Lib = eval(Fs.OpenTextFile(Folder + "accessdb.js", 1).ReadAll());

var JSON = {};
eval(Fs.OpenTextFile(Folder + "json2.js", 1).ReadAll());

if(typeof(ACCESSdb) == 'undefined' || typeof(JSON) == 'undefined'){
	WScript.Echo("ACCESSdb or JSON is not valid");
	WScript.Quit(1);
}
	
var confContent = Fs.OpenTextFile(Folder + "config.json", 1).ReadAll();
var Config = null;
eval("Config = " + confContent);
if(typeof(Config) == 'undefined' || Config == null){
	alert("Config is not valid");
	WScript.Quit(1);
}

var myDB = new ACCESSdb(Folder + "mdb.mdb", {showErrors:true});

function template2string(t, data){
	if(typeof(t) != 'string'){
		alert(t);
		return t;
	}
	var start = t.indexOf("{{");
	if(start < 0)
		return t;
	start = start + 2;
	var end = t.indexOf("}}", start);
	if(end < 0)
		return t;
	end = end + 2;
	//ok, we have variable
	t = t.substr(0, start - 2) + data[t.substr(start, end - start - 2)] + t.substr(end);
	return template2string(t, data);
}

function construct(data, config){
	var SQL = config.query;
	SQL = template2string(SQL, data);
	var context = myDB.query(SQL);
	
	var result = [];
	for(var c in context){
		if (!context.hasOwnProperty(c))
			continue;

		var obj = {};
		var ci = context[c];
		var t = [];
		for(var i in config.template){
			if (!config.template.hasOwnProperty(i))
				continue;
	
			var field = config.template[i];
			if(typeof(field) == 'object'){
				obj[i] = construct(ci, field);
			}else{
				obj[i] = template2string(field, ci);
			}
		}
		result.push(obj);
	}
	return result;
}

var r = construct({}, Config);
alert(JSON.stringify(r));