#target indesign

var idpp_location_parts = app.activeScript.fsName.split('\\');
idpp_location_parts.pop();
idpp_location_parts.push("TOMEI-KOKA.idpp");
idpp_location = idpp_location_parts.join('\\');




var pp = app.loadPreflightProfile(idpp_location);
var process = app.preflightProcesses.add(app.activeDocument, pp);
var links = app.activeDocument.links;
var files = new Array();
var pages = new Array();
var problem_messages = new Array();
var view_buttons = new Array();



var problems_string = "";

process.waitForProcess();

var results = process.processResults.replace(/(\r\n|\n|\r)/gm,"");
var split_results = results.split(" ");

if ((split_results[2] != "(0)") && (split_results.length > 3))
{
    number_of_errors = split_results[2].replace("(", "").replace("):", "");
    
    for (var r = 3; r < split_results.length; r++)
    {
        if (split_results[r] != "")
        {
            files.push(split_results[r])
            pages.push(split_results[r+1].replace("(R=", "").replace(")", ""));
            r+=2
        }
    }

    var box = new Window('dialog', "透明効果の情報");  
    box.size = [900, 400];
    box.margins = 5;
    box.alignment = 'top';
    box.title = box.add('statictext', undefined, "下記の透明効果があるオブジェクトが発見されました：");
    
    var main_group = box.add("group");
    var total_coverage_panel =main_group.add('panel', [0,0,800,200]);
        total_coverage_panel.margings = 5;
        total_coverage_panel.orientation = 'column';
        total_coverage_panel.alignment= 'fill';
    var scroller = main_group.add('scrollbar', [0,0,20,200]);    
    
    for (var f = 0; f < files.length; f++)
    {
        var one_group = total_coverage_panel.add('group');
        one_group.alignChildren = "right";
        problems_string = "ページ " + pages[f] + " にある " + files[f];
        problem_messages[f] = one_group.add('statictext', undefined, problems_string);
        view_buttons[f] = one_group.add('checkbox', undefined, 'ファイルを見る', { name: 'edit'});
    }

    box.button_panel = box.add('panel', undefined);
    box.button_panel.add('statictext', undefined, "Ok をクリックすると、「ファイルを見る」をチェックしたファイルの全てがエクスプローラーに開けます。");
    OK = box.button_panel.add('button', undefined, 'Ok', { name: 'ok'});
    box.show();
} 
else 
{
    alert("このファイルに --透明効果オブジェクト-- がありません。");
}

for (var f = 0; f < files.length; f++)
{
    if (view_buttons[f].value)
    {
        for (var l = 0; l < links.length; l++)
        {
            if (links[l].name == files[f])
            {
                links[l].revealInSystem();
            }
        }
    }
}

