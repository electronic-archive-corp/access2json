//force to use cscript and if arch is 64 bit - use correct cscript to start
(function(ws) {
	var WshShell = WScript.CreateObject ("WScript.Shell");
	var arch = WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%");
	var cmd = 'cscript.exe //nologo "' + ws.scriptFullName + '"';
	var correctArch = true;
	if(arch.indexOf("86") < 0){	
		var sr = WshShell.ExpandEnvironmentStrings("%SystemRoot%");
		cmd = sr + '\\SysWOW64\\' + cmd;
		correctArch = false;
	}
	var args = ws.arguments;
	for (var i = 0, len = args.length; i < len; i++) {
		var arg = args(i);
		cmd += ' ' + (~arg.indexOf(' ') ? '"' + arg + '"' : arg);
	}
	//cmd += ' 2>&1';
	
	function readAllFromAny(oExec){
		if (!oExec.StdOut.AtEndOfStream)
			return oExec.StdOut.ReadLine();

		if (!oExec.StdErr.AtEndOfStream)
			return "STDERR: " + oExec.StdErr.ReadLine();
		return -1;
	}

	// Execute a command line function....
	function callAndWait(execStr) {
		var oExec = WshShell.Exec(execStr);
		var TextStream = oExec.StdOut;
		while (oExec.Status == 0){
			while(!TextStream.AtEndOfStream){
				ws.Echo(TextStream.ReadLine());
			}
			//WScript.Sleep(10);
		}
	}

	if (!correctArch || ws.fullName.slice(-12).toLowerCase() !== '\\cscript.exe') {
		callAndWait(cmd);
		ws.quit();
	}
})(WScript);

var Fs = new ActiveXObject("Scripting.FileSystemObject");
var WshShell = WScript.CreateObject ("WScript.Shell");

var OutFile = null;

function log(str){
	var t;
	if(typeof(str) == 'undefined'){
		t = 'undefined';
	}else if(str === null){
		t = 'null';
	}else{
		t = str;
	}
	if(OutFile){
		OutFile.Write(t); 
	}else{
		WScript.StdOut.Write(t);
	}
}

function Quit(code){
	if(OutFile){
		OutFile.Close();
	}
}

// WshShell.CurrentDirectory is directory of process
// not directory of this file
var objFile = Fs.GetFile(WScript.ScriptFullName);
var Folder = Fs.GetParentFolderName(objFile) + "\\";
//log("Using folder: " + Folder);

var jsonFile = Folder + "json2.js";
if (!Fs.FileExists(jsonFile)){
	WScript.Echo("json2.js file not found");
	Quit(1);
}

var JSON = {};
eval(Fs.OpenTextFile(jsonFile, 1).ReadAll());

if(typeof(JSON) == 'undefined' || typeof(JSON.stringify) == 'undefined'){
	WScript.Echo("JSON object is not valid");
	Quit(1);
}

var dbPath = Folder + "mdb.mdb";
var confPath = Folder + "config.json";

var objArgs = WScript.Arguments;
/*
log("Args " + objArgs.length);
for(a in objArgs){
	log(a + ": " + objArgs[a]);
}
*/

if(objArgs.length >= 1){
	if(objArgs[0] != 'null'){
		dbPath = objArgs[0];
		log("using dbPath: " + dbPath);
	}
}

if(objArgs.length > 1){
	if(objArgs[1] != 'null'){
		confPath = objArgs[1];
	}
}

if(objArgs.length > 2){
	var f = Folder + "result.json";
	if(objArgs[2] == 'null'){
		f = objArgs[2];
	}
	OutFile = Fs.CreateTextFile(f, 8, true); //append
}

if (!Fs.FileExists(dbPath)){
	WScript.Echo("Database file " + dbPath + " not found");
	Quit(1);
}
	
if (!Fs.FileExists(confPath)){
	WScript.Echo("Config file " + confPath + " not found");
	Quit(1);
}

var confContent = Fs.OpenTextFile(confPath, 1).ReadAll();
var Config = null;
eval("Config = " + confContent);
if(typeof(Config) == 'undefined' || Config == null){
	log("Config is not valid");
	Quit(1);
}

function template2string(t, data, forQuery){
	function __jsonOrNot(t){
		if(forQuery)
			return t;
		return JSON.stringify(t);
	}
	if(typeof(t) != 'string'){
		return t;
	}
	var start = t.indexOf("{{");
	if(start < 0)
		return __jsonOrNot(t);
	start = start + 2;
	var end = t.indexOf("}}", start);
	if(end < 0)
		return __jsonOrNot(t);
	end = end + 2;
	//ok, we have variable
	var left = t.substr(0, start - 2);
	var right = t.substr(end);
	t =  data[t.substr(start, end - start - 2)];
	if (typeof(t) == "string") {
		t = t.replace(/\"/g, '&#34;');
		t = t.replace(/\'/g, '&#39;');
	}else if (typeof(t) == "date") {
		//t = "new Date(\"" + t + "\")";
	}else if (typeof(t) == "number") {
	}else{
		//debug ?
	}
	if(left.length == 0 && right.length == 0){
		return __jsonOrNot(t);
	}
	t = left + t + right;
	return __jsonOrNot(template2string(t, data, true));
}

function construct(config, data){
	var SQL = template2string(config.query, data, true);
	var rs = new ActiveXObject("ADODB.Recordset");
	var adOpenDynamic = 2;
	var adLockOptimistic = 3;
	rs.open(SQL, conn, adOpenDynamic, adLockOptimistic);
	log("[");
	if (rs.Fields.Count) {
		if (!rs.bof && !rs.eof){
			rs.MoveFirst();
			var first = true;
			while (!rs.eof){
				if(!first){
					log(",");
				}

				var dbRow = {};
				for (var x = 0; x < rs.Fields.Count; x++){
					dbRow[rs.Fields(x).Name] = rs.Fields(x).Value;
				}
				log('{');
				var firstProp = true;
				for(var i in config.template){
					if (!config.template.hasOwnProperty(i))
						continue;
					if(!firstProp)
						log(",");
					log('"' + i + '":')
					var field = config.template[i];
					if(typeof(field) == 'object'){
						construct(field, dbRow);
					}else{
						var val = template2string(field, dbRow, false);
						log(val);
					}
					firstProp = false;
				}
				log('}');
				first = false;
				rs.MoveNext();
			}
		}
	}
	log("]");
	rs.close();
}

var conn = null; // Connection to database

try{
	var cs = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + dbPath + ";Jet OLEDB:Engine Type=4;Persist Security Info = false"
	conn = new ActiveXObject("ADODB.Connection");
	conn.open(cs, "", ""); // connection string, user, pass
	construct(Config, {});
	conn.close();
}catch (e){
	log(e.name + "\n\n" + e.description);
	Quit(2);
}
Quit(0);