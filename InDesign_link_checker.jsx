#target InDesign

l_w_l_problems = new Array();
var l_w_l_problems_string = "";
l_w_l_problems_counter = 0;
l_problems = new Array();
var l_problems_string = "";
l_problems_counter = 0;

	// Load the XMP Script library
	if( xmpLib == undefined ) 
	{
		if( Folder.fs == "Windows" )
		{
			var pathToLib = Folder.startup.fsName + "/AdobeXMPScript.dll";
		} 
		else 
		{
			var pathToLib = Folder.startup.fsName + "/AdobeXMPScript.framework";
		}
	
		var libfile = new File( pathToLib );
		var xmpLib = new ExternalObject("lib:" + pathToLib );
	}
	
	// Get the selected file
	var thumb = File.openDialog("リンク情報を知りたいファイルを選んでください", undefined, false);
    if (thumb)
    {
    var base_name = thumb.name;
	
    
    xmpFile = new XMPFile(thumb.fsName, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
    var xmp = xmpFile.getXMP();
    
        var items = xmp.countArrayItems("http://ns.adobe.com/xap/1.0/mm/","Ingredients");
        if (items > 0)
        {
            for (var w = 1; w <= items; w++)
            {
                var property = "Ingredients[" + w + "]/stRef:filePath"
                var prop = xmp.getProperty("http://ns.adobe.com/xap/1.0/mm/",property);
                prop = prop.toString().substr(5,);
                fileObject = new File(prop);
                var prop_name = fileObject.name
                if (fileObject.exists)
                {
                    var thumb2 = new File(fileObject);
                    xmpFile2 =  new XMPFile(thumb2.fsName, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
                    var xmp2 = xmpFile2.getXMP();

                    var items2 = xmp2.countArrayItems("http://ns.adobe.com/xap/1.0/mm/","Manifest");
                    if (items2 > 0)
                    {
                        for (var b = 1; b <= items2; b++)
                        {
                            var property = "Manifest[" + b + "]/stMfs:reference/stRef:filePath"
                            var prop2 = xmp2.getProperty("http://ns.adobe.com/xap/1.0/mm/",property);
                            fileObject = new File(prop2);
                            prop2_name = fileObject.name;
                            if (fileObject.exists)
                            {
                                var nothing = 0;
                            }
                            else
                            {
                                var error_message = base_name + "   とリンクされている   " + prop_name + "   のリンクされているファイル  ( " + prop2_name + " ）  が存在しません.";
                                l_w_l_problems.push(error_message);
                            }
                        }
                    }
                }
                else
                {
                    var error_message = "" + base_name + " とリンクされていたファイルは、もう、存在していません。  ファイル名は： " + prop_name;
                    l_problems.push(error_message);
                }
            }
        }
    
        var problems = 0;
        var box = new Window('dialog', "リンクエラー情報");  
        box.size = [900, 800];
        box.margins = 5;
        box.alignment = 'top';
        box.title = box.add('statictext', undefined, "元のファイル ( " + base_name + " ) とリンクされているファイルの問題： (いわゆる「一層目リンク」の問題」)");
        box.one_layer_problems_panel = box.add('panel', undefined);
        box.one_layer_problems_panel.margins = 5;
        box.one_layer_problems_panel.orientation = 'row';
        box.one_layer_problems_panel.alignment = 'fill';
        if (l_problems.length === 0)
        {
            box.one_layer_problems_panel.add('statictext', undefined, '問題がありません。');
        }
        else
        {
            for (var b = 0; b <l_problems.length; b++)
            {
                l_problems_string = l_problems_string + " \n" +  l_problems[b];
            }
            st_links = box.one_layer_problems_panel.add('edittext', undefined, l_problems_string, {multiline:true});
            st_links.size = [900,300];
        }
        box.title = box.add('statictext', undefined, "" + base_name + "  とリンクされているファイルの中のリンク問題： (いわゆる「二層目リンク」の問題」)");
        box.two_layer_problems_panel = box.add('panel', undefined);
        box.two_layer_problems_panel.margins = 5;
        box.two_layer_problems_panel.orientation = 'row';
        box.two_layer_problems_panel.alignment = 'fill';
        if (l_w_l_problems.length === 0)
        {
            box.two_layer_problems_panel.add('statictext', undefined, '問題がありません。');
        }
        else
        {
            for (var a = 0; a < l_w_l_problems.length; a++)
            {
                l_w_l_problems_string = l_w_l_problems_string + " \n" + l_w_l_problems[a];
            }
            st_links2 = box.two_layer_problems_panel.add('edittext', undefined, l_w_l_problems_string, {multiline:true});
            st_links2.size = [900,300];
        }
        box.button_panel = box.add('panel', undefined);
        OK = box.button_panel.add('button', undefined, 'Ok', { name: 'ok'});
        box.show();

        
        

	}


