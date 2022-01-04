#target illustrator





alert("使用手順：\n\n1) *.ai 又は *.eps (又は両方）ファイルが入っているフォルダを選択します。\n2) フォルダ内の全てのファイルが加工されます。\n3) 加工済みの各ファイルは /Outlined というフォルダに保存されます。", "流れ", 0);

var sourceFolder = Folder.selectDialog( 'フォルダを選択してください');
var files = new Array ();

//app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

if (sourceFolder)
{
    files = sourceFolder.getFiles(/\.(eps|ai)$/i);

    var exportFolder = Folder(sourceFolder + "/Outlined");
    
    if(exportFolder.exists)
    {
        var removeFiles = exportFolder.getFiles();
        for (var a = 0; a < removeFiles.length; a++)
        {
            removeFiles[a].remove();
        }
        exportFolder.remove();
    }


    for (var j = 0; j < files.length; j++)
    {
        var optRef = new OpenOptions();
        optRef.updateLegacyText = true;
        var targetDocument = app.open(files[j], null, optRef);
        var tDLayers = targetDocument.layers;
        for (var L = 0; L < tDLayers.length; L++)
        {
            if (tDLayers[L].locked == false)
            {
                do
                {
                    for (var p = 0 ; p < tDLayers[L].textFrames.length; p++)
                    {
                        tDLayers[L].textFrames[p].createOutline(true);
                    }
                } while (tDLayers[L].textFrames.length > 0)
            }
        } 
        if(!exportFolder.exists)
        {
            exportFolder.create();
        }

        if (String(targetDocument.name).indexOf(".eps") > -1)
        {
            var saveName = new File (exportFolder + "/" + files[j].name);
            var saveOpts = new EPSSaveOptions();
            targetDocument.saveAs(saveName, saveOpts);
        }
        if (String(targetDocument.name).indexOf(".ai") > -1)
        {
            var saveName = new File (exportFolder + "/" + files[j].name);
            targetDocument.saveAs(saveName);
        }
        targetDocument.close();
    }
    alert("フォルダ内の全てのファイルのテキスト処理が終了です。\n\n「元の場所/Outlined」 のフォルダに保存されました。", "終了", 0);

}












