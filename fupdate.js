/* ================== Script Information Header =============================
 * Script Name:     fupdate.js
 * Date:            09.11.2021 �.
 * Author:          ������� ���� ����������
 * Description:     ������������� ����� ���������� ������ �� ����������
 *                  "fupdate.rar|zip|7z|arj" (� ��., ��� ���������� ������������
 *                  ��������� 7-zip) �� ��������� ������� Windows. ��������
 *                  ����� �� ������ �� ����, �������� ������-������ � ����
 *                  ������ � ������ � ���� "c|d|e|� �. �." (���� ������ �����
 *                  ���, �� ��-��������� ���������� ���� "d:"), � �����
 *                  ��������� �������, ���������� � ���� �����. ������� ������
 *                  ���� � ����������� "bat|cmd|py|js|vbs". ����������� �����
 *                  ��� ����� ����������� ������ ������ � ����������� �� �����
 *                  ����� ������� (��������: "0_.py" - �����, "_128.js" - �����).
 *                  ���������� ���������� ������� ���� �������� ������� ��
 *                  �������� ����� ����� ����� �������, �. �. "0_.cmd"
 *                  ���������� ������, ��� "1_.cmd".
 * ========================================================================== */

/* ================== Initialization Block ================================== */
// ���������
var MAX_SIZE_LOGFILE = 1048576; // ������������ ������ ���-����� - 1 �����
var SCRIPT_WIN_STYLE = 0;       // ����� ���� ��� ���������� �������
/** 
 * SCRIPT_WIN_STYLE = 0 -- ������ ����
 * SCRIPT_WIN_STYLE = 1 -- �������� ����
 **/

// �������
var fupdUnpack;      // ���������� ������
var fupdParse;       // ������ ����������� �������������� ������
var moveScripts;     // ����������� �������� �� ��������� �������
var runScript;       // ���������� �������
var createTempStore; // �������� ���������� ��������
var writeLog;        // �����������

// ������� ����������
var msgLog, fupdPath, fupd, prefix, postfix, file, folder;

// ����������� ��������
var args     = WScript.Arguments;
var wshShell = WScript.CreateObject("WScript.Shell");
var fso      = WScript.CreateObject("Scripting.FileSystemObject");

/* ================ Script Main Logic ======================================= */

msgLog = "----------------------";
writeLog(msgLog + ">> ������ ��������� <<" + msgLog);

try {
    if (fso.fileExists(args(0))) {
        writeLog(">>> ����������:");
        fupdPath = fupdUnpack(args(0));
        if (fupdPath) {
            fupd = fupdParse(fupdPath);

            if (fupd.prefixScripts.length || fupd.postfixScripts.length) {
                writeLog(">>> ����������� ��������:");
                if (fupd.prefixScripts.length) { // ����������
                    prefix = moveScripts(fupd.prefixScripts);
                }

                if (fupd.postfixScripts.length) { // �����������
                    postfix = moveScripts(fupd.postfixScripts);
                }
            }

            if (fupd.prefixScripts.length) {
                writeLog(">>> ���������� ���������� ��������:");
                for (var i=0; i<=prefix.scripts.length-1; i++) {
                    runScript(prefix.scripts[i], fupdPath);
                }
            }

            // ����������� �������� ������ � ���������.
            if (fupd.archFiles.length || fupd.archFolders.length) {
                writeLog(">>> ����������� ������ � ���������:");
                if (fupd.archFiles.length) {
                    for (var i=0; i<=fupd.archFiles.length-1; i++) {
                        file = fupd.archFiles[i];
                        msgLog  = "\t- ���� '";
                        msgLog += fso.GetFileName(file);
                        msgLog += "' ������� ���������� � �������";
                        file.Copy(fupd.driveLetter + "\\", true);
                        writeLog(msgLog);
                    }
                }

                if (fupd.archFolders.length) {
                    for (var i=0; i<=fupd.archFolders.length-1; i++) {
                        folder = fupd.archFolders[i];
                        msgLog  = "\t- ������� '";
                        msgLog += fso.GetFileName(folder);
                        msgLog += "' ������� ���������� � �������";
                        folder.Copy(fupd.driveLetter + "\\", true);
                        writeLog(msgLog);
                    }
                }
            }

            if (fupd.postfixScripts.length) {
                writeLog(">>> ���������� ����������� ��������:");
                for (var i=0; i<=postfix.scripts.length-1; i++) {
                    runScript(postfix.scripts[i], fupdPath);
                }
            }

            if (fupdPath || prefix || postfix) {
                writeLog(">>> �������:");
                if (fso.folderExists(fupdPath)) {
                    msgLog  = "\t- ��������� ������� '";
                    msgLog += fso.GetFileName(fupdPath);
                    msgLog += "' ������� �����";
                    fso.GetFolder(fupdPath).Delete(true);
                    writeLog(msgLog);
                }

                if (prefix) {
                    if (fso.folderExists(prefix.path)) {
                        msgLog  = "\t- ��������� ������� '";
                        msgLog += fso.GetFileName(prefix.path);
                        msgLog += "' ������� �����";
                        fso.GetFolder(prefix.path).Delete(true);
                        writeLog(msgLog);
                    }
                }

                if (postfix) {
                    if (fso.folderExists(postfix.path)) {
                        msgLog  = "\t- ��������� ������� '";
                        msgLog += fso.GetFileName(postfix.path);
                        msgLog += "' ������� �����";
                        fso.GetFolder(postfix.path).Delete(true);
                        writeLog(msgLog);
                    }
                }
            }
        }
    } else {
        writeLog("\t- ��� ��������� �����");
    }
}  catch (err) {
    /* ��������� ������ err */
    msgLog  = "\t- ������! - ";
    msgLog += err;
    writeLog(msgLog);
}

msgLog = "----------------------";
writeLog(msgLog + ">> �����  ��������� <<" + msgLog);
WScript.Quit(0);

/* ================ Functions =============================================== */

// ����������� ����������� ��������.
// ���������� ������, ���������� ������ ������������ �������� � ���� � ���.
function moveScripts(scripts) {
    var script, scripts, outList, newPath, msgLog;
    // ��������� ������� ����������
    var result = {
                    scripts : new Array(),
                    path    : createTempStore()
                 };
    for (var i=0; i<=scripts.length-1; i++) {
        script = scripts[i];
        script.Move(result.path + "\\");
        msgLog  = "\t- ������ '";
        msgLog += fso.GetFileName(script);
        msgLog += "' ��������� � ������� '";
        msgLog += fso.GetFileName(result.path);
        msgLog += "'";
        writeLog(msgLog);
        result.scripts.push(script);
    }
    return result;
}

// �������� ���������� �������� � ���������� C:\Windows\Temp.
// ���������� ���� � ������� ���������� ���������� ��������.
function createTempStore() {
    var msgLog, tempPath;
    tempPath  = wshShell.Environment.Item("TEMP");
    tempPath += "\\";
    tempPath += Math.random().toString(36).replace('0.', 'fupd_');
    tempPath  = wshShell.ExpandEnvironmentStrings(tempPath);

    if (fso.folderExists(tempPath)) {
        fso.GetFolder(tempPath).Delete(true);
    }

    msgLog  = "\t- ������ ��������� ������� '";
    msgLog += tempPath;
    msgLog += "'";
    fso.createFolder(tempPath);  // �������� ���������� ��������
    writeLog(msgLog);
    return tempPath;
}

// ���������� ������ fupdate.
// ���������� ���� � ������������� ������ ��� "null".
function fupdUnpack(archPath) {
    var archPath, unpackPath, runString, msgLog;

    unpackPath = createTempStore();

    runString  = "7z.exe x -r -aoa ";
    runString += archPath; // ���� � ������ fupdate
    runString += " -o";
    runString += unpackPath;

    if (!wshShell.Run(runString, 0, true)) {
        msgLog  = "\t- �� ��������� ������� '";
        msgLog += fso.GetFileName(unpackPath);
        msgLog += "' ���������� ����� '";
        msgLog += fso.GetFileName(archPath);
        msgLog += "'";
        writeLog(msgLog);
    } else {
        msgLog = "\t- ������! - ����� '";
        msgLog += fso.GetFileName(archPath);
        msgLog += "' �� ����������";
        writeLog(msgLog);
        msgLog = "\t- ��������� ������� '";
        msgLog += fso.GetFileName(unpackPath);
        msgLog += "' ������� �����";
        fso.GetFolder(unpackPath).Delete(true);
        writeLog(msgLog);
        unpackPath = null;
    }
    msgLog  = "\t- ����� '";
    msgLog += fso.GetFileName(archPath);
    msgLog += "' ������� �����";
    fso.GetFile(archPath).Delete(true);
    writeLog(msgLog);
    return unpackPath;
}

// ���������� ������ ����������� �������������� ������ fupdate.
// ���������� ������, ���������� ���� �����������, ������ ��������� ��������,
// ������ ���������� ��������, ������ �������� ������, ������ �������� ���������.
/**
 * �������!
 *      - ���������� ������� ����: "0_.bat", "1_.cmd", "258_.js", "840_.py"
 *                                      ����������� �� ����������� ������.
 *      - ����������� ������� ����: "_0.bat", "_1.cmd", "_258.js", "_840.py"
 *                                      ����������� ����� ����������� ������.
 * ������� ����������� � ������� ���������� �������� ����� ����� ����� �������
 * ��� ����������� �� ����������, �. �. ������ ���������� ������ "0_.bat" ("_0.bat")
 * �. �. � ������� ���������� ������. ����� ����� ���� �����, ���� "325847_.js"
 **/
function fupdParse(fupdPath) {
    var msgLog, file, files, folder, folders, fupdPath, fupdFolder;
    var driveLetter, prefixScripts, postfixScripts, archFiles, archFolders;

    var driveLetterPattern    = /^[A-Z]{1}$/i;
    var prefixScriptsPattern  = /^\d+_\.*/i;
    var postfixScriptsPattern = /^_\d+\.*/i;

    // ��������� ������� ����������
    var result = {
                    driveLetter    : "",
                    prefixScripts  : new Array(),
                    postfixScripts : new Array(),
                    archFiles      : new Array(),
                    archFolders    : new Array()
                 };

    writeLog(">>> ������ ������ ������ ����������:");

    fupdFolder = fso.GetFolder(fupdPath);

    files = new Enumerator(fupdFolder.Files);
    for (; !files.atEnd(); files.moveNext()) {
        file = files.item();
        if (driveLetterPattern.test(file.name)) {
            msgLog = "\t- ������� ����-����� '";
            msgLog += file.name;
            msgLog += "'";
            writeLog(msgLog);
            result.driveLetter = file.name.toLowerCase() + ":";
            if (!fso.DriveExists(result.driveLetter)) {
                msgLog = "\t- ���� '";
                msgLog += result.driveLetter;
                msgLog += "' �� ������!";
                writeLog(msgLog);
                result.driveLetter = "d:";
            }
            msgLog = "\t- ���������� ���� ��� ����������� � '";
            msgLog += result.driveLetter;
            msgLog += "'";
            writeLog(msgLog);
            msgLog = "\t- ������� ����-����� '";
            msgLog += file.name;
            msgLog += "'";
            fso.DeleteFile(file);
            writeLog(msgLog);
        } else if (prefixScriptsPattern.test(file.name)) {
            msgLog = "\t- ������ ���������� ������ '";
            msgLog += file.name;
            msgLog += "'";
            result.prefixScripts.push(file);
            writeLog(msgLog);
        } else if (postfixScriptsPattern.test(file.name)) {
            msgLog = "\t- ������ ����������� ������ '";
            msgLog += file.name;
            msgLog += "'";
            result.postfixScripts.push(file);
            writeLog(msgLog);
        } else {
            msgLog = "\t- ������ �������� ���� '";
            msgLog += file.name;
            msgLog += "'";
            result.archFiles.push(file);
            writeLog(msgLog);
        }
    }

    if (fupdFolder.SubFolders.Count) {
        folders = new Enumerator(fupdFolder.SubFolders);
        for (; !folders.atEnd(); folders.moveNext()) {
            folder = folders.item();
            msgLog = "\t- ������ �������� ������� '";
            msgLog += folder.name;
            msgLog += "'";
            result.archFolders.push(folder);
            writeLog(msgLog);
        }
    }

    if (!result.driveLetter) {
        writeLog("\t- ����-����� �� �������!");
        result.driveLetter = "d:";
        msgLog = "\t- ���������� ���� ��� ����������� � '";
        msgLog += result.driveLetter;
        msgLog += "'";
        writeLog(msgLog);
    }

    return result;
}

// ���������� ������� (����������� ��� ������������).
// ���������� ��������� ��������, 0 - ������ ������� ������� ��� ��� ������.
function runScript(scriptPath, archPath) {
    var scriptPath, archPath, runString, exitCode, msgLog;
    var runScriptData = {
                            "cmd" : "",
                            "bat" : "",
                            "js"  : "wscript.exe //nologo /e:jscript",
                            "vbs" : "wscript.exe //nologo /e:vbscript",
                            "py"  : "python.exe"
                        };
    runString  = runScriptData[fso.getExtensionName(scriptPath).toLowerCase()];
    runString += " ";
    runString += scriptPath;
    runString += " ";
    runString += archPath;
    runString  = runString.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');

    exitCode = wshShell.Run(runString, SCRIPT_WIN_STYLE, true);
    if (!exitCode) {
        msgLog  = "\t- ������ '";
        msgLog += fso.GetFileName(scriptPath);
        msgLog += "' �������� ������";
        writeLog(msgLog);
    } else {
        msgLog  = "\t- ������! - ������ '";
        msgLog += fso.GetFileName(scriptPath);
        msgLog += "' �� ��������";
        writeLog(msgLog);
    }
    return exitCode;
}

/* ================ Auxiliary Functions ===================================== */

// �����������.
function writeLog(msg) {
    var dt, year, month, day, hours, minutes, seconds;
    var msg, logFileName, logArchFileName, count, dtStamp;

    logFileName = fso.getBaseName(WScript.scriptName) + ".log";

    dt      = new Date();
    day     = zfill(dt.getDate(), 2);
    month   = zfill(dt.getMonth() + 1, 2);
    year    = dt.getYear();
    hours   = zfill(dt.getHours(), 2);
    minutes = zfill(dt.getMinutes(), 2);
    seconds = zfill(dt.getSeconds(), 2);

    dtStamp  = day + "." + month + "." + year;
    dtStamp += " ";
    dtStamp += hours + ":" + minutes + ":" + seconds;
    dtStamp += " | ";

    if (fso.fileExists(logFileName))
    {
        var file = fso.getFile(logFileName);
        if (file.size >= MAX_SIZE_LOGFILE)
        {
            count = 0;
            while (true)
            {
                logArchFileName = logFileName + "." + zfill(count, 3);
                if (fso.fileExists(logArchFileName))
                {
                    count++;
                }
                else
                {
                    file.move(logArchFileName);
                    break;
                }
            }
        }
    }

    var logFile = fso.openTextFile(logFileName, 8, true);
    logFile.writeLine(dtStamp + msg);
    logFile.close();
}

// ���������� ������ �������� ������ �� ����� n.
function zfill(input, n) {
    var input;
    var str = "";
    str += input;
    while(str.length < n)
    {
        str = "0" + str;
    }
    return str;
}
