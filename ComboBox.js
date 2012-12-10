function ShowList() {
	var temp = event.srcElement.id.split("_");
	if(document.getElementById(temp[0] + "_DivList").style.visibility == "hidden") {
		document.getElementById(temp[0] + "_DivList").style.visibility = "visible";
		document.getElementById(temp[0] + "_listBox").focus();
	} else {
		document.getElementById(temp[0] + "_DivList").style.visibility = "hidden";
	}
}

function FillText() {
	var temp = event.srcElement.id.split("_");
	var select = document.getElementById(event.srcElement.id);
	if (select.selectedIndex != -1)
		document.getElementById(temp[0] + "_Texte").value = select[select.selectedIndex].text;
	
	document.getElementById(temp[0] + "_DivList").style.visibility = "hidden";
}

function HideList() {
	if (event != null) {
		var evt = event.srcElement.id.split("_");
		
		tables = document.getElementsByTagName("DIV");
			
		for (var i = 0; i < tables.length; i++) {
			var Cache = true;
			var temp = tables[i].id.split("_")
			if(temp.length == 2) {
				if(evt != null) {
					if(evt.length = 2) {
						if(evt[1] == "Arrow" && temp[0] == evt[0]) {
							Cache = false;	
						}
					}
				}
				if(Cache) {
					if(temp[1] == "DivList") {
						document.getElementById(tables[i].id).style.visibility = "hidden";
					} 
				}
			}			
		} 
	}
}

function TextChange() {
	var temp = event.srcElement.id.split("_");
	var Liste = document.getElementById(temp[0] + "_listBox");
	var text = document.getElementById(temp[0] + "_Texte");
	var valeur = text.value;
	
	var trouve = false;	
	var g = -1;
	var d = Liste.length;
	
	if(event.keyCode >= 32 && !(event.keyCode >= 37 && event.keyCode <= 39)) {
		while(d-g > 1 && !trouve) {
			var m = (g+d)/2;
			var pos = -1;
			var pos=Liste[m].text.toUpperCase().indexOf(valeur.toUpperCase());
			
			if(pos == 0) {
				trouve=true;
			} 
			if(valeur.toUpperCase() < Liste[m].text.toUpperCase()) {
				d = m;
			} else {
				g = m;
			}
		}
		
		if(trouve) {
			while(m > 0 && Liste[m].text.toUpperCase().indexOf(valeur.toUpperCase()) == 0) {
				m--;
			}
			
			Liste.selectedIndex = m+1;
			text.value = valeur + Liste[m + 1].text.substr(valeur.length, Liste[m + 1].text.length - valeur.length);
			var oRange = text.createTextRange(); 
			oRange.moveStart("character", valeur.length); 
			oRange.moveEnd("character", Liste[m + 1].text.length - valeur.length);      
			oRange.select();
		}
	}
	document.getElementById(temp[0] + "_TextisChanged").value = 1;
}

function PostBack() {
	var temp = event.srcElement.id.split("_");
	var changer = document.getElementById(temp[0] + "_TextisChanged");
	if(changer.value == 1) {
		changer.value = 0;
		document.forms[0].submit();
	}
}



