var postSubmited = false;
function ctlent(obj) {
	if(postSubmited == false && (event.ctrlKey && window.event.keyCode == 13) || (event.altKey && window.event.keyCode == 83)) {
			if(this.document.input.topicsubmit) this.document.input.topicsubmit.disabled = true;
			if(this.document.input.replysubmit) this.document.input.replysubmit.disabled = true;
			if(this.document.input.editsubmit) this.document.input.editsubmit.disabled = true;
			postSubmited = true;
			this.document.input.submit();
	}
}

var flag=false; 
function DrawImage(ImgD){ 
	var image=new Image(); 
	image.src=ImgD.src; 
	if(image.width>0 && image.height>0){ 
		flag=true; 
		if(image.width>=518){ 
			ImgD.width=518; 
			ImgD.height=(image.height*518)/image.width; 
		}else{ 
			ImgD.width=image.width; 
			ImgD.height=image.height; 
		}  
	} 
} 

function showIntro(objID)
{
	if (document.getElementById(objID).style.display == "none") {
		document.getElementById(objID).style.display = "";
	}else{
		document.getElementById(objID).style.display = "none";
	}
}

function winconfirm(confirmMSG,confirmURL){
	question = confirm(confirmMSG);
	if (question != "0"){
		window.location=confirmURL;
	}
}

function CopyText(obj) {
	ie = (document.all)? true:false
	if (ie){
		var rng = document.body.createTextRange();
		rng.moveToElementText(obj);
		rng.scrollIntoView();
		rng.select();
		rng.execCommand("Copy");
		rng.collapse(false);
	}
}

function UBBShowObj(strType,strID,strURL,intWidth,intHeight)
{
	var varHeader="V";
	var tmpstr="";
	var bSwitch = false;
	bSwitch = document.getElementById(varHeader+strID).value;
	bSwitch	=~bSwitch;
	document.getElementById(varHeader+strID).value = bSwitch;
	if(bSwitch){
		document.getElementById(strID).innerHTML = "影音文件源地址：<a href='"+strURL+"' target='_blank'>"+strURL+"</a>";
	}else{
		switch(strType.toUpperCase()){
			case "SWF":
				tmpstr="<object codeBase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0' classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' width='"+intWidth+"' height='"+intHeight+"'><param name='movie' value='"+strURL+"'><param name='quality' value='high'><param name='AllowScriptAccess' value='never'><embed src='"+strURL+"' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='"+intWidth+"' height='"+intHeight+"'>'"+strURL+"'</embed></OBJECT>";
				break;
			case "WMA":
				tmpstr="<object classid='CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95'  id='MediaPlayer' width='"+intWidth+"' height='"+intHeight+"'><param name='ShowStatusBar' value='-1'><param name='AutoStart' value='True'><param name='Filename' value='"+strURL+"'></object>";
				break;
			case "WMV":
				tmpstr="<object classid='CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95' width='"+intWidth+"' height='"+intHeight+"'><param name='ShowStatusBar' value='-1'><param name='AutoStart' value='Ture'><param name='Filename' value='"+strURL+"'></object>";
				break;
			case "RM":
				tmpstr="<object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' height='"+intHeight+"' id='Player' width='"+intWidth+"' viewastext><param name='_ExtentX' value='12726'><param name='_ExtentY' value='8520'><param name='AUTOSTART' value='0'><param name='SHUFFLE' value='0'><param name='PREFETCH' value='0'><param name='NOLABELS' value='0'><param name='CONTROLS' value='ImageWindow'><param name='CONSOLE' value='_master'><param name='LOOP' value='0'><param name='NUMLOOP' value='0'><param name='CENTER' value='0'><param name='MAINTAINASPECT' value='"+strURL+"'><param name='BACKGROUNDCOLOR' value='#000000'></object><br><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' height='32' id='Player' width='"+intWidth+"' VIEWASTEXT><param name='_ExtentX' value='18256'><param name='_ExtentY' value='794'><param name='AUTOSTART' value='-1'><param name='SHUFFLE' value='0'><param name='PREFETCH' value='0'><param name='NOLABELS' value='0'><param name='CONTROLS' value='controlpanel'><param name='CONSOLE' value='_master'><param name='LOOP' value='0'><param name='NUMLOOP' value='0'><param name='CENTER' value='0'><param name='MAINTAINASPECT' value='0'><param name='BACKGROUNDCOLOR' value='#000000'><param name='SRC' value='"+strURL+"'></object>";
				break;
			case "QT":
				tmpstr="<embed src='"+strURL+"' autoplay='true' Loop='false' controller='true' playeveryframe='false' cache='false' scale='TOFIT' bgcolor='#000000' kioskmode='false' targetcache='false' pluginspage='http://www.apple.com/quicktime/'>";
				break;
			case "RA":
				tmpstr="<object classid='clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA' id='RAOCX' width='"+intWidth+"' height='"+intHeight+"'><param name='_ExtentX' value='6694'><param name='_ExtentY' value='1588'><param name='AUTOSTART' value='0'><param name='SHUFFLE' value='0'><param name='PREFETCH' value='0'><param name='NOLABELS' value='0'><param name='SRC' value='"+strURL+"'><param name='CONTROLS' value='StatusBar,ControlPanel'><param name='LOOP' value='0'><param name='NUMLOOP' value='0'><param name='CENTER' value='0'><param name='MAINTAINASPECT' value='0'><param name='BACKGROUNDCOLOR' value='#000000'><embed src='"+strURL+"' width='253' autostart='true' height='60'></embed></object>"
		}
		document.getElementById(strID).innerHTML = tmpstr;
	}
}

function AutoSizeDIV(objID){
	var obj=document.getElementById(objID);
	if (obj.scrollHeight < obj.offsetHeight){
		obj.style.height = obj.scrollHeight+2;
	}
}

//-------------背单词------------
var En=new Array()
var Ch=new Array()
var Ei=0
function showE(first){
if (first!=1) SEng.filters[0].apply();
SEng.innerHTML="<img src='images/zd.gif' align='absmiddle'> <span class=EnText>"+En[Ei]+"</span><div class=ChText>"+Ch[Ei]+"</div>"
if (first!=1) SEng.filters[0].play();
Ei++
if (Ei==En.length) {Ei=0}
}