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
function alert(str){
	WScript.Echo(str);
}

//alert(arch);

// this one is directory of process
//alert (WshShell.CurrentDirectory);

//find current directory of this file
var objFile = Fs.GetFile(WScript.ScriptFullName);
var Folder = Fs.GetParentFolderName(objFile) + "\\";
//alert("Using folder: " + Folder);

var jsonFile = Folder + "json2.js";
if (!Fs.FileExists(jsonFile)){
	WScript.Echo("json2.js file not found");
	WScript.Quit(1);
}

var JSON = {};
eval(Fs.OpenTextFile(jsonFile, 1).ReadAll());

if(typeof(JSON) == 'undefined' || typeof(JSON.stringify) == 'undefined'){
	WScript.Echo("JSON object is not valid");
	WScript.Quit(1);
}

var dbPath = Folder + "mdb.mdb";
var confPath = Folder + "config.json";

var objArgs = WScript.Arguments;
//alert("Args " + objArgs.length);
for(a in objArgs){
	alert(a + ": " + objArgs[a]);
}
if(objArgs.length >= 1)
{
   dbPath = objArgs[0];
   alert("using dbPath: " + dbPath);
}

if(objArgs.length > 1)
{
   confPath = objArgs[1];
}

if (!Fs.FileExists(dbPath)){
	WScript.Echo("Database file " + dbPath + " not found");
	WScript.Quit(1);
}
	
if (!Fs.FileExists(confPath)){
	WScript.Echo("Config file " + confPath + " not found");
	WScript.Quit(1);
}

var confContent = Fs.OpenTextFile(confPath, 1).ReadAll();
var Config = null;
eval("Config = " + confContent);
if(typeof(Config) == 'undefined' || Config == null){
	alert("Config is not valid");
	WScript.Quit(1);
}

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
	var SQL = template2string(config.query, data);
	var context = query(SQL);
	
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

var conn = null; // Connection to database

function query(q){
	try{
		var rs = new ActiveXObject("ADODB.Recordset");
		var adOpenDynamic = 2;
		var adLockOptimistic = 3;
		rs.open(q, conn, adOpenDynamic, adLockOptimistic);
		if (rs.Fields.Count) {
			if (!rs.bof && !rs.eof){
				var r = [];
				rs.MoveFirst();
				while (!rs.eof){
					var obj = {};
					for (var x = 0; x < rs.Fields.Count; x++){
						obj[rs.Fields(x).Name] = rs.Fields(x).Value;
					}
					rs.MoveNext();
					r.push(obj);
				}
				return r;
			}
		}
		rs.close();
	}catch (e){
		alert("Query " + e.name + "\n\n" + e.description);
	}
	return false;
}

try{
	//Connection string
	var cs = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + dbPath + ";Jet OLEDB:Engine Type=4;Persist Security Info = false"
	//Microsoft.ACE.OLEDB.12.0
	conn = new ActiveXObject("ADODB.Connection");
	conn.open(cs, "", ""); // connection string, user, pass
	var r = construct({}, Config);
	alert(JSON.stringify(r));
	conn.close();
}catch (e){
	alert(e.name + "\n\n" + e.description);
}