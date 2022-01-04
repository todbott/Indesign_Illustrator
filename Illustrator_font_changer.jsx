#target illustrator

/*
    フォント交換　（Illustrator CS5)
    Ver 2.1
    2018/03/01
    ギリース・トッド
    
    Ver 2.0 の変更されたところは、フォントの選択仕方。複数のターゲットフォントを選択する後、複数の置きたいフォントも選択できるようになりました。
    Ver 2.1 の変更されたところは、*.eps 対応を加えました。
    
   */



var sourceFolder = Folder.selectDialog( 'フォントの変えたいファイルはどこにありますか？フォルダを選択してください');
var files = new Array ();
var found_font_array = new Array ();　//holds all found fonts names, full of duplicates
var found_family_array = new Array (); // holds the found font family names
var found_style_array = new Array (); // holds the found font style names

var system_font_array = new Array ();　//holds all found fonts names, full of duplicates
var system_family_array = new Array (); // holds the found font family names
var system_style_array = new Array (); // holds the found font style names

var system_font_array_2 = new Array ();　//holds all found fonts names, full of duplicates
var system_family_array_2 = new Array (); // holds the found font family names
var system_style_array_2 = new Array (); // holds the found font style names

var found_font_array_2 = new Array (); //holds all found font names, with no duplicates
var found_family_array_2 = new Array (); // holds the found font family names with no duplicates
var found_style_array_2 = new Array (); // holds the found font style names with no duplicates
var ddown_array_family = new Array (); //the dropdown box list of available fonts in the system by family name
var ddown_array_style = new Array (); //the dropdown box list of available fonts in the system by style
var targeted_font_string_array = new Array (); //array full of fonts the user wants to target
var replacement_font_string_array = new Array (); //array full of fonts the user wants to put in
var empty = []; //an empty array for use later

var targeted_font_string_array_for_error_code = new Array ();

var selected;

var relation_array = new Array;

for (var f = 0; f < app.textFonts.length; f++)
{
    var font = String(app.textFonts[f].name)
    relation_array.push(font)
    var family_and_style = String(app.textFonts[f].family) + " " + String(app.textFonts[f].style)
    relation_array.push(family_and_style)
}

//******************************************************************************************************//
//	対象ファイルの中に、存在しているフォントのフォント名を把握する
//******************************************************************************************************//

app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;


if (sourceFolder)
{
    files = sourceFolder.getFiles(/\.(eps|ai)$/i);

    var exportFolder = Folder(sourceFolder + "/Henko_sumi_ai");
    var finalFolder = Folder(sourceFolder + "/Henko_sumi_eps");
    
    if(exportFolder.exists)
    {
        var removeFiles = exportFolder.getFiles();
        for (var a = 0; a < removeFiles.length; a++)
        {
            removeFiles[a].remove();
        }
        exportFolder.remove();
    }
    if(finalFolder.exists)
    {
        var removeFiles = finalFolder.getFiles();
        for (var a = 0; a < removeFiles.length; a++)
        {
            removeFiles[a].remove();
        }
        finalFolder.remove();
    }

var already_discovered = "";
var o = 0;

for (var j = 0; j < files.length; j++)
{
    var optRef = new OpenOptions();
    optRef.updateLegacyText = true;
    var targetDocument = app.open(files[j], DocumentColorSpace.CMYK, optRef);
    if (targetDocument.textFrames.length > 0)
    {
        for (var p = 0 ; p < targetDocument.textFrames.length; p++)
        {
            textArtRange = targetDocument.textFrames[p].textRange;
            var i = 0;
            do
            {
                try
                {
                    var found_font = textArtRange.characterAttributes.textFont.name;
                    var found_font_family = textArtRange.characterAttributes.textFont.family;
                    var found_font_style = textArtRange.characterAttributes.textFont.style;
                }
                catch (e)
                {
                    if (e instanceof TypeError)
                    {
                        i = i + 1;
                        textArtRange = targetDocument.textFrames[i].textRange;
                        var error_holder = e;
                    }
                }
                if (error_holder == "")
                {
                    break;
                }
                error_holder = "";
            }
            while (i < 1000) 
            if (already_discovered.indexOf(found_font + " ") < 0)
            {
                found_font_array[o] = found_font;
                found_family_array[o] = found_font_family;
                found_style_array[o] = found_font_style;
                o = o + 1
            }
            already_discovered = already_discovered + "," + found_font + " ";
        }
    }
    if (String(targetDocument.name).indexOf(".eps") > -1)
    {
        if(!exportFolder.exists)
        {
            exportFolder.create();
        }
        if(!finalFolder.exists)
        {
            finalFolder.create();
        }
        var saveName = new File (exportFolder + "/" + files[j].name);
        targetDocument.saveAs(saveName);
    }
    if (String(targetDocument.name).indexOf(".ai") > -1)
    {
        var saveName = new File (targetDocument.path + "/" + files[j].name);
        targetDocument.saveAs(saveName);
    }
    targetDocument.close(SaveOptions.DONOTSAVECHANGES);
}


//******************************************************************************************************//
//	remove duplicates from system fonts
//******************************************************************************************************//

already_discovered = "";
var found_font;
var found_family;
var found_style;
for (var b = 0; b < app.textFonts.length; b++)
{
    found_font = app.textFonts[b];
    found_family = app.textFonts[b].family;
    found_style = app.textFonts[b].style;
    if (already_discovered.indexOf(found_font) < 0)
            {
                system_font_array[b] = found_font;
            }
            if (already_discovered.indexOf(found_family) < 0)
            {
                system_family_array[b] = found_family;
            }
            if (already_discovered.indexOf(found_style) < 0)
            {
                system_style_array[b] = found_style;
            }
    already_discovered = already_discovered + "," + app.textFonts[b] + "," + app.textFonts[b].family + "," + app.textFonts[b].style;
}
    
for (var ff = 0;  ff < system_font_array.length; ff++)
{
    if (system_font_array[ff])
    {
        system_font_array_2.push(system_font_array[ff]);
    }
}

for (var ff = 0;  ff < system_family_array.length; ff++)
{
    if (system_family_array[ff])
    {
        system_family_array_2.push(system_family_array[ff]);
    }
}

for (var ff = 0;  ff < system_style_array.length; ff++)
{
    if (system_style_array[ff])
    {
        system_style_array_2.push(system_style_array[ff]);
    }
}

//******************************************************************************************************//
//	remove duplicates from found fonts
//******************************************************************************************************//
for (var ff = 0;  ff < found_font_array.length; ff++)
{
    if (found_font_array[ff])
    {
        found_font_array_2.push(found_font_array[ff]);
        found_family_array_2.push(found_family_array[ff]);
        found_style_array_2.push(found_style_array[ff]);
    }
}





//******************************************************************************************************//
//	ユーザーにダイアログボックスを見せる。
//******************************************************************************************************//

var box = new Window('dialog', "　"　+ found_font_array_2.length + " フォントが発見しました");  
var popup = new Window('dialog', " ");

box.title = box.add('statictext', undefined, 'このファイルの内に見つけたフォントの代わりに、どちらのフォントを入りたいですか？');
box.title = box.add('statictext', undefined, '「変えない」を選択するとすれば、変えないフォントを設定することも出来ます。');
box.size = [600, 600];
box.margins = 5;
box.alignment = 'top';

for (var a = 0; a < found_font_array_2.length; a++)
{
    box.found_fonts_panel = box.add('panel', undefined);
    box.found_fonts_panel.margins = 5;
    box.found_fonts_panel.orientation = 'row';
    box.found_fonts_panel.alignment = 'fill';
    box.found_fonts_panel.add('statictext', undefined, '   発見した　' + found_family_array_2[a] + ' ' + found_style_array_2[a] + '   を  ');
    ddown_array_family[a] = box.found_fonts_panel.add('dropdownlist', undefined, empty);
    ddown_array_style[a] = box.found_fonts_panel.add('dropdownlist', undefined, empty);
    ddown_array_style[a].size = [100, 20];
    box.found_fonts_panel.add('statictext', undefined, '  に変えます');
    ddown_array_family[a].add('item','変えない');
    for (var p = 0; p < system_family_array_2.length; p++)
    {
        ddown_array_family[a].add('item', system_family_array_2[p]);
    }
    ddown_array_family[a].selection = 0;
}



var funcs = [];

function createfunc(p) {
  return function() {
        ddown_array_style[p].removeAll();
        for (var w = 0; w < app.textFonts.length; w++)
        {
            selected = String(ddown_array_family[p].selection);
            in_system = String(app.textFonts[w].family)
            if (selected === in_system)
            {
                ddown_array_style[p].add('item', app.textFonts[w].style);
            }
        };
    }
}


for (var p = 0; p < found_font_array_2.length; p++)
{
    funcs[p] = createfunc(p);
}

for (var p = 0; p < found_font_array_2.length; p++)
{
    ddown_array_family[p].onChange = funcs[p];
}



box.button_panel = box.add('panel', undefined, "フォントを入り変わります");  
OK = box.button_panel.add('button', undefined, 'Ok', { name: 'ok'});
CANCEL = box.button_panel.add('button', undefined, 'Cancel', { name: 'キャンセル'});



//******************************************************************************************************//
//	選択されたフォントとターゲットフォントを　string として array に入れる。
//******************************************************************************************************//

OK.onClick = function change_fonts () {
    var problem_font_list = "";
    var position = 0;
    for (var p = 0; p < found_font_array_2.length; p++)
    {
        if (String(ddown_array_family[p].selection) != "変えない")
        {
            replacement_font_string_array[p] = String(ddown_array_family[p].selection) + " " + String(ddown_array_style[p].selection);
            var temp = replacement_font_string_array[p];
            for (var r = 0; r < relation_array.length; r++)
            {
                if (relation_array[r] === temp)
                {
                    position = r;
                    position = position - 1;
                }
            }
        replacement_font_string_array[p] = String(relation_array[position]);
        }
        else
        {
            replacement_font_string_array[p] = "変えない";
        }
        targeted_font_string_array[p] = String(found_font_array_2[p]);
        targeted_font_string_array_for_error_code[p] = String(found_family_array_2[p]);
    }

//******************************************************************************************************//
//	フォントを変えて、保存して、閉じる。
//******************************************************************************************************//

    if (exportFolder)
    {
        files = exportFolder.getFiles("*.ai");
    
        for (var j = 0; j < files.length; j++)
        {
            var targetDocument = app.open(files[j]);
            if (targetDocument.textFrames.length > 0)
            {
                for (var i = 0 ; i < targetDocument.textFrames.length; i++)
                {
                    textArtRange = targetDocument.textFrames[i].textRange;
                    do
                    {
                        try
                        {
                            var found_font = textArtRange.characterAttributes.textFont.name;
                        }
                        catch (e)
                        {
                            if (e instanceof TypeError)
                            {
                                i = i + 1;
                                textArtRange = targetDocument.textFrames[i].textRange;
                                var error_holder = e;
                            }
                        }
                        if (error_holder == "")
                        {
                            break;
                        }
                        error_holder = "";
                    }
                    while (i < 10000)
                
                    var found_font = String(textArtRange.characterAttributes.textFont.name);
                
                    for (var p = 0; p < found_font_array_2.length; p++)
                    {
                        var replacement_font_string = replacement_font_string_array[p];
                        var targeted_font_string = targeted_font_string_array[p];
                        if ((replacement_font_string != "変えない") && (found_font === targeted_font_string))
                        { 
                            var should_be_replaced = String(textArtRange.characterAttributes.textFont);
                            textArtRange.characterAttributes.textFont = textFonts.getByName(replacement_font_string);
                            var shouldve_been_replaced = String(textArtRange.characterAttributes.textFont);
                            if (should_be_replaced === shouldve_been_replaced)
                            {
                                if (problem_font_list.indexOf(targeted_font_string_array_for_error_code[p]) < 0)
                                {
                                    problem_font_list = targeted_font_string_array_for_error_code[p] + "   " + problem_font_list
                                }  
                            }
                        }
                    }
                }
            }
        var saveName = new File (finalFolder + "/" + targetDocument.name);
        var saveOpts = new EPSSaveOptions();
        saveOpts.cmykPostScript = true;
        saveOpts.embedAllFonts = true;
        targetDocument.saveAs(saveName, saveOpts);
        
        var saveName2 = new File (exportFolder + "/" + targetDocument.name);
        targetDocument.saveAs(saveName2);
        targetDocument.close(SaveOptions.DONOTSAVECHANGES);
        }
    }

    files = sourceFolder.getFiles("*.ai")
    if (files.length > 0)
    {
        for (var j = 0; j < files.length; j++)
        {
            var targetDocument = app.open(files[j]);
            if (targetDocument.textFrames.length > 0)
            {
                for (var i = 0 ; i < targetDocument.textFrames.length; i++)
                {
                    textArtRange = targetDocument.textFrames[i].textRange;
                    do
                    {
                        try
                        {
                            var found_font = textArtRange.characterAttributes.textFont.name;
                        }
                        catch (e)
                        {
                        if (e instanceof TypeError)
                            {
                                i = i + 1;
                                textArtRange = targetDocument.textFrames[i].textRange;
                                var error_holder = e;
                            }
                        }
                        if (error_holder == "")
                        {
                            break;
                        }
                        error_holder = "";
                    }
                    while (i < 10000)
              
                    var found_font = String(textArtRange.characterAttributes.textFont.name);
               
                    for (var p = 0; p < found_font_array_2.length; p++)
                    {
                        var replacement_font_string = replacement_font_string_array[p];
                        var targeted_font_string = targeted_font_string_array[p];
                        if ((replacement_font_string != "変えない") && (found_font === targeted_font_string))
                        { 
                            var should_be_replaced = String(textArtRange.characterAttributes.textFont);
                            textArtRange.characterAttributes.textFont = textFonts.getByName(replacement_font_string);
                            var shouldve_been_replaced = String(textArtRange.characterAttributes.textFont);
                            if (should_be_replaced === shouldve_been_replaced)
                            {
                                if (problem_font_list.indexOf(targeted_font_string_array_for_error_code[p]) < 0)
                                {
                                    problem_font_list = targeted_font_string_array_for_error_code[p] + "   " + problem_font_list
                                }  
                            }
                        }
                    }
                }
            }
            targetDocument.saveAs(targetDocument.fullName);
            targetDocument.close(SaveOptions.DONOTSAVECHANGES);
        }
    }

    
    alert("終わりました。\n\n先加工した*.eps ファイル（ある場合）は、「Henko_sumi_eps」フォルダに保存されていて、*.ai バージョンは 「Henko_sumi_ai」 フォルダに保存されています。\n\n先加工した*.ai ファイル（ある場合）は、元の場所に保存されています。");
    if (problem_font_list != "")
    {
        alert("次のフォントは、設定されたフォントに交換できませんでした： \n" + problem_font_list + "。　\n\n場合によりますが、フォントファミリが違えば、交換が出来ない可能性があります。\n\n(例えば、「KS]から始まるフォントが、「KS」から始まるフォント以外に交換できません。)");
    }
    box.close();
};
CANCEL.onClick = function c () {
    alert("キャンセルされました");
    box.close();
};
box.show() ;
}
else
{
    alert("キャンセルされました");
}
