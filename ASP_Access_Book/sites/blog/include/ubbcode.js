defmode = "normalmode";		// default mode (normalmode, advmode, helpmode)

if (defmode == "advmode") {
        helpmode = false;
        normalmode = false;
        advmode = true;
} else if (defmode == "helpmode") {
        helpmode = true;
        normalmode = false;
        advmode = false;
} else {
        helpmode = false;
        normalmode = true;
        advmode = false;
}
function chmode(swtch){
        if (swtch == 1){
                advmode = false;
                normalmode = false;
                helpmode = true;
                alert(help_mode);
        } else if (swtch == 0) {
                helpmode = false;
                normalmode = false;
                advmode = true;
                alert(adv_mode);
        } else if (swtch == 2) {
                helpmode = false;
                advmode = false;
                normalmode = true;
                alert(normal_mode);
        }
}

function AddText(NewCode) {
        if(document.all){
        	insertAtCaret(document.input.message, NewCode);
        	setfocus();
        } else{
        	document.input.message.value += NewCode;
        	setfocus();
        }
}

function storeCaret (textEl){
        if(textEl.createTextRange){
                textEl.caretPos = document.selection.createRange().duplicate();
        }
}

function insertAtCaret (textEl, text){
        if (textEl.createTextRange && textEl.caretPos){
                var caretPos = textEl.caretPos;
                caretPos.text += caretPos.text.charAt(caretPos.text.length - 2) == ' ' ? text + ' ' : text;
        } else if(textEl) {
                textEl.value += text;
        } else {
        	textEl.value = text;
        }
}

function email() {
        if (helpmode) {
                alert(email_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[email]" + range.text + "[/email]";
	} else if (advmode) {
	      AddTxt="[email] [/email]";
                AddText(AddTxt);
        } else { 
                txt2=prompt(email_normal,""); 
                if (txt2!=null) {
                        txt=prompt(email_normal_input,"name@domain.com");      
                        if (txt!=null) {
                                if (txt2=="") {
                                        AddTxt="[email]"+txt+"[/email]";
                
                                } else {
                                        AddTxt="[email="+txt+"]"+txt2+"[/email]";
                                } 
                                AddText(AddTxt);                
                        }
                }
        }
}


function chsize(size) {
        if (helpmode) {
                alert(fontsize_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[size=" + size + "]" + range.text + "[/size]";
        } else if (advmode) {
                AddTxt="[size="+size+"] [/size]";
                AddText(AddTxt);
        } else {                       
                txt=prompt(fontsize_normal,text_input); 
                if (txt!=null) {             
                        AddTxt="[size="+size+"]"+txt;
                        AddText(AddTxt);
                        AddText("[/size]");
                }        
        }
}

function chfont(font) {
        if (helpmode){
                alert(font_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[font=" + font + "]" + range.text + "[/font]";
        } else if (advmode) {
                AddTxt="[font="+font+"] [/font]";
                AddText(AddTxt);
        } else {                  
                txt=prompt(font_normal,text_input);
                if (txt!=null) {             
                        AddTxt="[font="+font+"]"+txt;
                        AddText(AddTxt);
                        AddText("[/font]");
                }        
        }  
}


function bold() {
        if (helpmode) {
                alert(bold_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[b]" + range.text + "[/b]";
        } else if (advmode) {
                AddTxt="[b] [/b]";
                AddText(AddTxt);
        } else {  
                txt=prompt(bold_normal,text_input);     
                if (txt!=null) {           
                        AddTxt="[b]"+txt;
                        AddText(AddTxt);
                        AddText("[/b]");
                }       
        }
}

function italicize() {
        if (helpmode) {
                alert(italicize_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[i]" + range.text + "[/i]";
        } else if (advmode) {
                AddTxt="[i] [/i]";
                AddText(AddTxt);
        } else {   
                txt=prompt(italicize_normal,text_input);     
                if (txt!=null) {           
                        AddTxt="[i]"+txt;
                        AddText(AddTxt);
                        AddText("[/i]");
                }               
        }
}

function quote() {
        if (helpmode){
                alert(quote_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[quote]" + range.text + "[/quote]";
        } else if (advmode) {
                AddTxt="\r[quote]\r[/quote]";
                AddText(AddTxt);
        } else {   
                txt=prompt(quote_normal,text_input);     
                if(txt!=null) {          
                        AddTxt="\r[quote]\r"+txt;
                        AddText(AddTxt);
                        AddText("\r[/quote]");
                }               
        }
}

function chcolor(color) {
        if (helpmode) {
                alert(color_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[color=" + color + "]" + range.text + "[/color]";
        } else if (advmode) {
                AddTxt="[color="+color+"] [/color]";
                AddText(AddTxt);
        } else {  
        txt=prompt(color_normal,text_input);
                if(txt!=null) {
                        AddTxt="[color="+color+"]"+txt;
                        AddText(AddTxt);
                        AddText("[/color]");
                }
        }
}

function center() {
        if (helpmode) {
                alert(center_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[align=center]" + range.text + "[/align]";
        } else if (advmode) {
                AddTxt="[align=center] [/align]";
                AddText(AddTxt);
        } else {  
                txt=prompt(center_normal,text_input);     
                if (txt!=null) {          
                        AddTxt="\r[align=center]"+txt;
                        AddText(AddTxt);
                        AddText("[/align]");
                }              
        }
}

function hyperlink() {
        if (helpmode) {
                alert(link_help);
        } else if (advmode) {
                AddTxt="[url] [/url]";
                AddText(AddTxt);
        } else { 
                txt2=prompt(link_normal,""); 
                if (txt2!=null) {
                        txt=prompt(link_normal_input,"http://");      
                        if (txt!=null) {
                                if (txt2=="") {
                                        AddTxt="[url]"+txt;
                                        AddText(AddTxt);
                                        AddText("[/url]");
                                } else {
                                        AddTxt="[url="+txt+"]"+txt2;
                                        AddText(AddTxt);
                                        AddText("[/url]");
                                }         
                        } 
                }
        }
}

function image() {
        if (helpmode){
                alert(image_help);
        } else if (advmode) {
                AddTxt="[img] [/img]";
                AddText(AddTxt);
        } else {  
                txt=prompt(image_normal,"http://");    
                if(txt!=null) {            
                        AddTxt="\r[img]"+txt;
                        AddText(AddTxt);
                        AddText("[/img]");
                }       
        }
}

function flash() {
        if (helpmode){
                alert(flash_help);
        } else if (advmode) {
                AddTxt="[swf] [/swf]";
                AddText(AddTxt);
        } else {  
                txt=prompt(flash_normal,"N"); 
				while ((txt!="") && (txt!="n") && (txt!="N") && (txt!="y") && (txt!="Y") && (txt!=null)) {
                        txt=prompt(flash_normal_error,""); 
				}
                if (txt!=null) {
                        if (txt=="") {
                                AddTxt="\r[swf]";
                        } else {
				if ((txt=="y") || (txt=="Y")) {
                                AddTxt="\r[swf=550,400,AP]";
				} else {
                                AddTxt="\r[swf]";
				}
                        } 
                        txt="n";
                        if(txt!=null) {
                                txt=prompt(flash_normal_input,"http://"); 
                                if (txt!="") {             
                                        AddTxt+=txt; 
                                }                   
                        } 
                        AddTxt+="[/swf]";
                        AddText(AddTxt);  
                }
        }
}

function code() {
        if (helpmode) {
                alert(code_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[code]" + range.text + "[/code]";
        } else if (advmode) {
                AddTxt="\r[code]\r[/code]";
                AddText(AddTxt);
        } else {   
                txt=prompt(code_normal,"");     
                if (txt!=null) {          
                        AddTxt="\r[code]"+txt;
                        AddText(AddTxt);
                        AddText("[/code]");
                }              
        }
}

function wma() {
        if (helpmode){
                alert(wma_help);
        } else if (advmode) {
                AddTxt="[wma] [/wma]";
                AddText(AddTxt);
        } else {  
                txt=prompt(wma_normal,"N"); 
				while ((txt!="") && (txt!="n") && (txt!="N") && (txt!="y") && (txt!="Y") && (txt!=null)) {
                        txt=prompt(wma_normal_error,""); 
				}
                if (txt!=null) {
                        if (txt=="") {
                                AddTxt="\r[wma]";
                        } else {
				if ((txt=="y") || (txt=="Y")) {
                                AddTxt="\r[wma=450,70,AP]";
				} else {
                                AddTxt="\r[wma]";
				}
                        } 
                        txt="n";
                        if(txt!=null) {
                                txt=prompt(wma_normal_input,"http://"); 
                                if (txt!="") {             
                                        AddTxt+=txt; 
                                }                   
                        } 
                        AddTxt+="[/wma]";
                        AddText(AddTxt);  
                }
        }
}

function wmv() {
        if (helpmode){
                alert(wmv_help);
        } else if (advmode) {
                AddTxt="[wmv] [/wmv]";
                AddText(AddTxt);
        } else {  
                txt=prompt(wmv_normal,"N"); 
				while ((txt!="") && (txt!="n") && (txt!="N") && (txt!="y") && (txt!="Y") && (txt!=null)) {
                        txt=prompt(wmv_normal_error,""); 
				}
                if (txt!=null) {
                        if (txt=="") {
                                AddTxt="\r[wmv]";
                        } else {
				if ((txt=="y") || (txt=="Y")) {
                                AddTxt="\r[wmv=550,400,AP]";
				} else {
                                AddTxt="\r[wmv]";
				}
                        } 
                        txt="n";
                        if(txt!=null) {
                                txt=prompt(wmv_normal_input,"http://"); 
                                if (txt!="") {             
                                        AddTxt+=txt; 
                                }                   
                        } 
                        AddTxt+="[/wmv]";
                        AddText(AddTxt);  
                }
        }
}

function ra() {
        if (helpmode){
                alert(ra_help);
        } else if (advmode) {
                AddTxt="[ra] [/ra]";
                AddText(AddTxt);
        } else {  
                txt=prompt(ra_normal,"N"); 
				while ((txt!="") && (txt!="n") && (txt!="N") && (txt!="y") && (txt!="Y") && (txt!=null)) {
                        txt=prompt(ra_normal_error,""); 
				}
                if (txt!=null) {
                        if (txt=="") {
                                AddTxt="\r[ra]";
                        } else {
				if ((txt=="y") || (txt=="Y")) {
                                AddTxt="\r[ra=450,60,AP]";
				} else {
                                AddTxt="\r[ra]";
				}
                        } 
                        txt="n";
                        if(txt!=null) {
                                txt=prompt(ra_normal_input,"http://"); 
                                if (txt!="") {             
                                        AddTxt+=txt; 
                                }                   
                        } 
                        AddTxt+="[/ra]";
                        AddText(AddTxt);  
                }
        }
}

function rm() {
        if (helpmode){
                alert(rm_help);
        } else if (advmode) {
                AddTxt="[rm] [/rm]";
                AddText(AddTxt);
        } else {  
                txt=prompt(rm_normal,"N"); 
				while ((txt!="") && (txt!="n") && (txt!="N") && (txt!="y") && (txt!="Y") && (txt!=null)) {
                        txt=prompt(rm_normal_error,""); 
				}
                if (txt!=null) {
                        if (txt=="") {
                                AddTxt="\r[rm]";
                        } else {
				if ((txt=="y") || (txt=="Y")) {
                                AddTxt="\r[rm=550,400,AP]";
				} else {
                                AddTxt="\r[rm]";
				}
                        } 
                        txt="n";
                        if(txt!=null) {
                                txt=prompt(rm_normal_input,"http://"); 
                                if (txt!="") {             
                                        AddTxt+=txt; 
                                }                   
                        } 
                        AddTxt+="[/rm]";
                        AddText(AddTxt);  
                }
        }
}

function page() {
        if (helpmode) {
                alert(page_help);
    } else if (document.selection && document.selection.type == "Text") {
        var range = document.selection.createRange();
        range.text = "[page]" + range.text + "[/page]";
        } else if (advmode) {
                AddTxt="[page] [/page]";
                AddText(AddTxt);
        } else {  
                txt=prompt(page_normal,text_input);     
                if (txt!=null) {           
                        AddTxt="[page]"+txt;
                        AddText(AddTxt);
                        AddText("[/page]");
                }       
        }
}

function light() {
        if (helpmode) {
                alert(light_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[light]" + range.text + "[/light]";
        } else if (advmode) {
                AddTxt="[light] [/light]";
                AddText(AddTxt);
        } else {  
                txt=prompt(light_normal,text_input);     
                if (txt!=null) {           
                        AddTxt="[light]"+txt;
                        AddText(AddTxt);
                        AddText("[/light]");
                }       
        }
}

function mem() {
        if (helpmode) {
                alert(mem_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[mem]" + range.text + "[/mem]";
        } else if (advmode) {
                AddTxt="[mem] [/mem]";
                AddText(AddTxt);
        } else {  
                txt=prompt(mem_normal,text_input);     
                if (txt!=null) {           
                        AddTxt="[mem]"+txt;
                        AddText(AddTxt);
                        AddText("[/mem]");
                }       
        }
}

function list() {
        if (helpmode) {
                alert(list_help);
        } else if (advmode) {
                AddTxt="\r[list]\r[*]\r[*]\r[*]\r[/list]";
                AddText(AddTxt);
        } else {  
                txt=prompt(list_normal,"");
                while ((txt!="") && (txt!="A") && (txt!="a") && (txt!="1") && (txt!=null)) {
                        txt=prompt(list_normal_error,"");               
                }
                if (txt!=null) {
                        if (txt=="") {
                                AddTxt="\r[list]\r\n";
                        } else {
                                AddTxt="\r[list="+txt+"]\r";
                        } 
                        txt="1";
                        while ((txt!="") && (txt!=null)) {
                                txt=prompt(list_normal_input,""); 
                                if (txt!="") {             
                                        AddTxt+="[*]"+txt+"\r"; 
                                }                   
                        } 
                        AddTxt+="[/list]\r\n";
                        AddText(AddTxt); 
                }
        }
}

function underline() {
        if (helpmode) {
                alert(underline_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[u]" + range.text + "[/u]";
        } else if (advmode) {
                AddTxt="[u] [/u]";
                AddText(AddTxt);
        } else {  
                txt=prompt(underline_normal,text_input);
                if (txt!=null) {           
                        AddTxt="[u]"+txt;
                        AddText(AddTxt);
                        AddText("[/u]");
                }               
        }
}

function sub() {
        if (helpmode) {
                alert(sub_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[sub]" + range.text + "[/sub]";
        } else if (advmode) {
                AddTxt="[sub] [/sub]";
                AddText(AddTxt);
        } else {  
                txt=prompt(sub_normal,text_input);
                if (txt!=null) {           
                        AddTxt="[sub]"+txt;
                        AddText(AddTxt);
                        AddText("[/sub]");
                }               
        }
}

function sup() {
        if (helpmode) {
                alert(sup_help);
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[sup]" + range.text + "[/sup]";
        } else if (advmode) {
                AddTxt="[sup] [/sup]";
                AddText(AddTxt);
        } else {  
                txt=prompt(sup_normal,text_input);
                if (txt!=null) {           
                        AddTxt="[sup]"+txt;
                        AddText(AddTxt);
                        AddText("[/sup]");
                }               
        }
}

function Cfly() {
        if (helpmode) {
                alert("飞行文字，例如\n[fly]test[/fly]");
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[fly]" + range.text + "[/fly]";
        } else if (advmode) {
                AddTxt="[fly] [/fly]";
                AddText(AddTxt);
        } else {  
                txt=prompt("请输入文字",text_input);
                if (txt!=null) {           
                        AddTxt="[fly]"+txt;
                        AddText(AddTxt);
                        AddText("[/fly]");
                }               
        }
}


function Cmove() {
        if (helpmode) {
                alert("移动文字，例如\n[move]test[/move]");
	} else if (document.selection && document.selection.type == "Text") {
		var range = document.selection.createRange();
		range.text = "[move]" + range.text + "[/move]";
        } else if (advmode) {
                AddTxt="[move] [/move]";
                AddText(AddTxt);
        } else {  
                txt=prompt("请输入文字",text_input);
                if (txt!=null) {           
                        AddTxt="[move]"+txt;
                        AddText(AddTxt);
                        AddText("[/move]");
                }               
        }
}

function Cglow() {
        if (helpmode) {
                alert("发光文字，例如\n[glow=255,red,2]test[/glow]");
        } else if (advmode) {
                AddTxt="[glow] [/glow]";
                AddText(AddTxt);
        } else { 
                txt2=prompt("文字的长度、颜色和边界大小","255,blue,2"); 
                if (txt2!=null) {
                        txt=prompt("请输入文字","文字");      
                        if (txt!=null) {
                                if (txt2=="") {
                                        AddTxt="[glow]"+txt;
                                        AddText(AddTxt);
                                        AddText("[/glow]");
                                } else {
                                        AddTxt="[glow="+txt2+"]"+txt;
                                        AddText(AddTxt);
                                        AddText("[/glow]");
                                }         
                        } 
                }
        }
}

function Cshadow() {
        if (helpmode) {
                alert("阴影文字，例如\n[glow=255,red,2]test[/glow]");
        } else if (advmode) {
                AddTxt="[shadow] [/shadow]";
                AddText(AddTxt);
        } else { 
                txt2=prompt("文字的长度、颜色和边界大小","255,blue,1"); 
                if (txt2!=null) {
                        txt=prompt("请输入文字","文字");      
                        if (txt!=null) {
                                if (txt2=="") {
                                        AddTxt="[shadow]"+txt;
                                        AddText(AddTxt);
                                        AddText("[/shadow]");
                                } else {
                                        AddTxt="[shadow="+txt2+"]"+txt;
                                        AddText(AddTxt);
                                        AddText("[/shadow]");
                                }         
                        } 
                }
        }
}

function download() {
        if (helpmode) {
                alert(download_help);
        } else if (advmode) {
                AddTxt="[download] [/download]";
                AddText(AddTxt);
        } else { 
                txt2=prompt(download_normal,"会员下载"); 
                if (txt2!=null) {
                        txt=prompt(download_normal_input,"http://");      
                        if (txt!=null) {
                                if (txt2=="") {
                                        AddTxt="[download]"+txt;
                                        AddText(AddTxt);
                                        AddText("[/download]");
                                } else {
                                        AddTxt="[download="+txt+"]"+txt2;
                                        AddText(AddTxt);
                                        AddText("[/download]");
                                }         
                        } 
                }
        }
}

function blogquote(objID,strAuthor,strTime){
	document.input.message.value += "[quote][b]最初由 [color=blue]"+strAuthor+"[/color] 发表于 "+strTime+"：[/b]\n"+document.getElementById(objID).innerText+"[/quote]";
	setfocus();
}

function setfocus() {
        document.input.message.focus();
}