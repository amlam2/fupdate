/* ================== Script Information Header =============================
 * Script Name:     fupdate.js
 * Date:            09.11.2021 г.
 * Author:          Лобович Олег Михайлович
 * Description:     Распаковывает архив обновлений файлов на компьютере
 *                  "fupdate.rar|zip|7z|arj" (и др., при распаковке используется
 *                  архиватор 7-zip) во временный каталог Windows. Копирует
 *                  файлы из архива на диск, заданный файлом-меткой в этом
 *                  архиве с именем в вида "c|d|e|и т. д." (если такого файла
 *                  нет, то по-умолчанию выбирается диск "d:"), а также
 *                  выполняет скрипты, помещённые в этот архив. Скрипты должны
 *                  быть с расширением "bat|cmd|py|js|vbs". Выполняются перед
 *                  или после копирования файлов архива в зависимости от имени
 *                  файла скрипта (например: "0_.py" - перед, "_128.js" - после).
 *                  Очерёдность выполнения каждого вида скриптов зависит от
 *                  цифровой части имени файла скрипта, т. е. "0_.cmd"
 *                  выполнится раньше, чем "1_.cmd".
 * ========================================================================== */

/* ================== Initialization Block ================================== */
// Константы
var MAX_SIZE_LOGFILE = 1048576; // максимальный размер лог-файла - 1 МБайт
var SCRIPT_WIN_STYLE = 0;       // стиль окна при выполнении скрипта
/** 
 * SCRIPT_WIN_STYLE = 0 -- скрыть окно
 * SCRIPT_WIN_STYLE = 1 -- показать окно
 **/

// Функции
var fupdUnpack;      // распаковка архива
var fupdParse;       // анализ содержимого распакованного архива
var moveScripts;     // перемещение скриптов во временный каталог
var runScript;       // выполнение скрипта
var createTempStore; // создание временного каталога
var writeLog;        // логирование

// Текущие переменные
var msgLog, fupdPath, fupd, prefix, postfix, file, folder;

// Подключение объектов
var args     = WScript.Arguments;
var wshShell = WScript.CreateObject("WScript.Shell");
var fso      = WScript.CreateObject("Scripting.FileSystemObject");

/* ================ Script Main Logic ======================================= */

msgLog = "----------------------";
writeLog(msgLog + ">> Начало обработки <<" + msgLog);

try {
    if (fso.fileExists(args(0))) {
        writeLog(">>> Распаковка:");
        fupdPath = fupdUnpack(args(0));
        if (fupdPath) {
            fupd = fupdParse(fupdPath);

            if (fupd.prefixScripts.length || fupd.postfixScripts.length) {
                writeLog(">>> Перемещение скриптов:");
                if (fupd.prefixScripts.length) { // префиксных
                    prefix = moveScripts(fupd.prefixScripts);
                }

                if (fupd.postfixScripts.length) { // постфиксных
                    postfix = moveScripts(fupd.postfixScripts);
                }
            }

            if (fupd.prefixScripts.length) {
                writeLog(">>> Выполнение префиксных скриптов:");
                for (var i=0; i<=prefix.scripts.length-1; i++) {
                    runScript(prefix.scripts[i], fupdPath);
                }
            }

            // Копирование архивных файлов и каталогов.
            if (fupd.archFiles.length || fupd.archFolders.length) {
                writeLog(">>> Копирование файлов и каталогов:");
                if (fupd.archFiles.length) {
                    for (var i=0; i<=fupd.archFiles.length-1; i++) {
                        file = fupd.archFiles[i];
                        msgLog  = "\t- файл '";
                        msgLog += fso.GetFileName(file);
                        msgLog += "' успешно скопирован с заменой";
                        file.Copy(fupd.driveLetter + "\\", true);
                        writeLog(msgLog);
                    }
                }

                if (fupd.archFolders.length) {
                    for (var i=0; i<=fupd.archFolders.length-1; i++) {
                        folder = fupd.archFolders[i];
                        msgLog  = "\t- каталог '";
                        msgLog += fso.GetFileName(folder);
                        msgLog += "' успешно скопирован с заменой";
                        folder.Copy(fupd.driveLetter + "\\", true);
                        writeLog(msgLog);
                    }
                }
            }

            if (fupd.postfixScripts.length) {
                writeLog(">>> Выполнение постфиксных скриптов:");
                for (var i=0; i<=postfix.scripts.length-1; i++) {
                    runScript(postfix.scripts[i], fupdPath);
                }
            }

            if (fupdPath || prefix || postfix) {
                writeLog(">>> Очистка:");
                if (fso.folderExists(fupdPath)) {
                    msgLog  = "\t- временный каталог '";
                    msgLog += fso.GetFileName(fupdPath);
                    msgLog += "' успешно удалён";
                    fso.GetFolder(fupdPath).Delete(true);
                    writeLog(msgLog);
                }

                if (prefix) {
                    if (fso.folderExists(prefix.path)) {
                        msgLog  = "\t- временный каталог '";
                        msgLog += fso.GetFileName(prefix.path);
                        msgLog += "' успешно удалён";
                        fso.GetFolder(prefix.path).Delete(true);
                        writeLog(msgLog);
                    }
                }

                if (postfix) {
                    if (fso.folderExists(postfix.path)) {
                        msgLog  = "\t- временный каталог '";
                        msgLog += fso.GetFileName(postfix.path);
                        msgLog += "' успешно удалён";
                        fso.GetFolder(postfix.path).Delete(true);
                        writeLog(msgLog);
                    }
                }
            }
        }
    } else {
        writeLog("\t- нет архивного файла");
    }
}  catch (err) {
    /* обработка ошибки err */
    msgLog  = "\t- ОШИБКА! - ";
    msgLog += err;
    writeLog(msgLog);
}

msgLog = "----------------------";
writeLog(msgLog + ">> Конец  обработки <<" + msgLog);
WScript.Quit(0);

/* ================ Functions =============================================== */

// Перемещение исполняемых скриптов.
// Возвращает объект, содержащий список перемещённых скриптов и путь к ним.
function moveScripts(scripts) {
    var script, scripts, outList, newPath, msgLog;
    // объектный литерал результата
    var result = {
                    scripts : new Array(),
                    path    : createTempStore()
                 };
    for (var i=0; i<=scripts.length-1; i++) {
        script = scripts[i];
        script.Move(result.path + "\\");
        msgLog  = "\t- скрипт '";
        msgLog += fso.GetFileName(script);
        msgLog += "' перемещён в каталог '";
        msgLog += fso.GetFileName(result.path);
        msgLog += "'";
        writeLog(msgLog);
        result.scripts.push(script);
    }
    return result;
}

// Создание временного каталога в директории C:\Windows\Temp.
// Возвращает путь к недавно созданному временному каталогу.
function createTempStore() {
    var msgLog, tempPath;
    tempPath  = wshShell.Environment.Item("TEMP");
    tempPath += "\\";
    tempPath += Math.random().toString(36).replace('0.', 'fupd_');
    tempPath  = wshShell.ExpandEnvironmentStrings(tempPath);

    if (fso.folderExists(tempPath)) {
        fso.GetFolder(tempPath).Delete(true);
    }

    msgLog  = "\t- создан временный каталог '";
    msgLog += tempPath;
    msgLog += "'";
    fso.createFolder(tempPath);  // создание временного каталога
    writeLog(msgLog);
    return tempPath;
}

// Распаковка архива fupdate.
// Возвращает путь к распакованным файлам или "null".
function fupdUnpack(archPath) {
    var archPath, unpackPath, runString, msgLog;

    unpackPath = createTempStore();

    runString  = "7z.exe x -r -aoa ";
    runString += archPath; // путь к архиву fupdate
    runString += " -o";
    runString += unpackPath;

    if (!wshShell.Run(runString, 0, true)) {
        msgLog  = "\t- во временный каталог '";
        msgLog += fso.GetFileName(unpackPath);
        msgLog += "' распакован архив '";
        msgLog += fso.GetFileName(archPath);
        msgLog += "'";
        writeLog(msgLog);
    } else {
        msgLog = "\t- ОШИБКА! - архив '";
        msgLog += fso.GetFileName(archPath);
        msgLog += "' не распакован";
        writeLog(msgLog);
        msgLog = "\t- временный каталог '";
        msgLog += fso.GetFileName(unpackPath);
        msgLog += "' успешно удалён";
        fso.GetFolder(unpackPath).Delete(true);
        writeLog(msgLog);
        unpackPath = null;
    }
    msgLog  = "\t- архив '";
    msgLog += fso.GetFileName(archPath);
    msgLog += "' успешно удалён";
    fso.GetFile(archPath).Delete(true);
    writeLog(msgLog);
    return unpackPath;
}

// Производит анализ содержимого распакованного архива fupdate.
// Возвращает объект, содержащий путь копирования, список префисных скриптов,
// список постфисных скриптов, список архивных файлов, список архивных каталогов.
/**
 * Справка!
 *      - префиксные скрипты вида: "0_.bat", "1_.cmd", "258_.js", "840_.py"
 *                                      Выполняются до копирования файлов.
 *      - постфиксные скрипты вида: "_0.bat", "_1.cmd", "_258.js", "_840.py"
 *                                      Выполняются после копирования файлов.
 * Скрипты выполняются в порядке увеличения цифровой части имени файла скрипта
 * вне зависимости от расширения, т. е. первым выполнится скрипт "0_.bat" ("_0.bat")
 * и. т. в порядке увеличения номера. Номер может быть любым, хоть "325847_.js"
 **/
function fupdParse(fupdPath) {
    var msgLog, file, files, folder, folders, fupdPath, fupdFolder;
    var driveLetter, prefixScripts, postfixScripts, archFiles, archFolders;

    var driveLetterPattern    = /^[A-Z]{1}$/i;
    var prefixScriptsPattern  = /^\d+_\.*/i;
    var postfixScriptsPattern = /^_\d+\.*/i;

    // объектный литерал результата
    var result = {
                    driveLetter    : "",
                    prefixScripts  : new Array(),
                    postfixScripts : new Array(),
                    archFiles      : new Array(),
                    archFolders    : new Array()
                 };

    writeLog(">>> Анализ файлов архива обновлений:");

    fupdFolder = fso.GetFolder(fupdPath);

    files = new Enumerator(fupdFolder.Files);
    for (; !files.atEnd(); files.moveNext()) {
        file = files.item();
        if (driveLetterPattern.test(file.name)) {
            msgLog = "\t- найдена файл-метка '";
            msgLog += file.name;
            msgLog += "'";
            writeLog(msgLog);
            result.driveLetter = file.name.toLowerCase() + ":";
            if (!fso.DriveExists(result.driveLetter)) {
                msgLog = "\t- диск '";
                msgLog += result.driveLetter;
                msgLog += "' не найден!";
                writeLog(msgLog);
                result.driveLetter = "d:";
            }
            msgLog = "\t- установлен диск для копирования в '";
            msgLog += result.driveLetter;
            msgLog += "'";
            writeLog(msgLog);
            msgLog = "\t- удалена файл-метка '";
            msgLog += file.name;
            msgLog += "'";
            fso.DeleteFile(file);
            writeLog(msgLog);
        } else if (prefixScriptsPattern.test(file.name)) {
            msgLog = "\t- найден префиксный скрипт '";
            msgLog += file.name;
            msgLog += "'";
            result.prefixScripts.push(file);
            writeLog(msgLog);
        } else if (postfixScriptsPattern.test(file.name)) {
            msgLog = "\t- найден постфиксный скрипт '";
            msgLog += file.name;
            msgLog += "'";
            result.postfixScripts.push(file);
            writeLog(msgLog);
        } else {
            msgLog = "\t- найден архивный файл '";
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
            msgLog = "\t- найден архивный каталог '";
            msgLog += folder.name;
            msgLog += "'";
            result.archFolders.push(folder);
            writeLog(msgLog);
        }
    }

    if (!result.driveLetter) {
        writeLog("\t- файл-метка не найдена!");
        result.driveLetter = "d:";
        msgLog = "\t- установлен диск для копирования в '";
        msgLog += result.driveLetter;
        msgLog += "'";
        writeLog(msgLog);
    }

    return result;
}

// Выполнение скрипта (префиксного или постфиксного).
// Возвращает результат операции, 0 - скрипт выполен успешно или код ошибки.
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
        msgLog  = "\t- скрипт '";
        msgLog += fso.GetFileName(scriptPath);
        msgLog += "' выполнен успешо";
        writeLog(msgLog);
    } else {
        msgLog  = "\t- ОШИБКА! - скрипт '";
        msgLog += fso.GetFileName(scriptPath);
        msgLog += "' не выполнен";
        writeLog(msgLog);
    }
    return exitCode;
}

/* ================ Auxiliary Functions ===================================== */

// Логирование.
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

// Дополнение строки ведущими нулями до длины n.
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
