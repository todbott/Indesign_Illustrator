#target indesign

function dialog () {
    var box = new Window('dialog', "Overflow対策");  

    box.export_panel = box.add('panel', undefined);
    
    box.export_panel.orientation = 'row';
    box.export_panel.alignment = 'fill';
    box.export_panel.add('statictext', undefined, '長体の最低限：');
    mc = box.export_panel.add('edittext', undefined, "50");
    box.export_panel.add('statictext', undefined, 'ポイント数の最低限：');
    mp = box.export_panel.add('edittext', undefined, "12"); 
    go_button = box.export_panel.add('button', undefined, 'GO!', {name: 'export'});
    close = box.export_panel.add('button', undefined, 'キャンセル', {name: 'cancel'});
    

    box.size = [750, 50];
    box.margins = 5;
    box.alignment = 'top';
    
    box.close_panel = box.add('panel', undefined);
    
    

    go_button.onClick = function () {
        box.close();
        }


    box.show();
}


dialog();

var doc = app.activeDocument
var items = doc.allPageItems

var min_chotai = parseInt(mc.text);
var current_chotai = 0;

var min_point = parseInt(mp.text);
var current_point = 0;

for (var i = 0; i < items.length; i++) 
{

    if (items[i].getElements()[0].constructor.name == "TextFrame") 
    {
        var tf = items[i].getElements()[0]; // get one text frame
        
         do
        {
            var tfp = tf.paragraphs; // get the text in that text frame

            for (var t = 0; t < tfp.length; t++) // for each paragraph in that text frame
            {
                tfp[t].minimumGlyphScaling = min_chotai;  // set the minimum chotai
                current_chotai = tfp[t].desiredGlyphScaling - 1;  // get the next chotai value to be implemented
                tfp[t].desiredGlyphScaling = current_chotai; // implement the chotai scaling
            }
            if ((current_point <= min_point) && (current_chotai <= min_chotai)) { break; }
            if (current_chotai <= min_chotai)
            {
                for (var t = 0; t < tfp.length; t++) // for each paragraph in that text frame
                {
                    current_point = tfp[t].pointSize -1;
                    tfp[t].pointSize = current_point;
                    current_chotai = 100;  
                    tfp[t].desiredGlyphScaling = current_chotai; 
                }
            }
        } while (tf.overflows)
    }
}

