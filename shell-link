function getStartMenuDirectory() {
    var shell = new ActiveXObject("WScript.Shell");
    var allUserStartMenu = shell.SpecialFolders("AllUsersStartMenu"),
        allUserStartMenuPrograms = shell.SpecialFolders("AllUsersPrograms");
    var currentUserStartMenu = shell.SpecialFolders("StartMenu"),
        currentUserStartMenuPrograms = shell.SpecialFolders("Programs");

    return allUserProgramsMenu;
}


function createWhere(where, source, filename) {
    var shell = new ActiveXObject("WScript.Shell");

    if (where.indexOf('cur') === 0) {
        where = shell.SpecialFolders("StartMenu");
    } else if (where.indexOf('all') === 0) {
        where = shell.SpecialFolders("AllUsersStartMenu");
    }

    if (!filename) {
        var slashPos = (source.lastIndexOf('\\') || source.lastIndexOf('/')) + 1;
            dotPos = source.lastIndexOf('.');

        if (dotPos <= slashPos) {
            dotPos = source.length;
        }

        filename = source.substr(slashPos, dotPos-slashPos);
    }

    WScript.echo('Creating in ' + where);
    WScript.echo('Filename is ' + where + '\\' + filename + '.lnk');

    var shortcut = shell.CreateShortcut(where + '\\' + filename + '.lnk');

    shortcut.TargetPath = source;
    //shortcut.Arguments = "\""+target_script_folder+"\\script.js\" /aParam /orTwo";
    shortcut.IconLocation = source;
    shortcut.Save();
}

if (WScript.arguments.length < 2) {
    WScript.echo('Usage: ' + WScript.ScriptName + ' currentUserMenu|allUserMenu|path_where_to_create_link source_path [target_link_filename]')
} else {
    createWhere(WScript.arguments(0), WScript.arguments(1), WScript.arguments.length > 2 ? WScript.arguments(2) : undefined);
}
