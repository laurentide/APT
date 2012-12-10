function ForecastChange() {
	var temp = event.srcElement.id.split("_");
	var NoCase = temp[1];

	if(document.forms[0].ForecastChanged.value.indexOf("{" + NoCase + "}") < 0) {
		document.forms[0].ForecastChanged.value += "{" +  NoCase + "};";
	}
}

function InitiativesChange() {
	var temp = event.srcElement.id.split("_");
	var NoCase = temp[1];

	if(document.forms[0].InitiativeChanged.value.indexOf("{" + NoCase + "}") < 0) {
		document.forms[0].InitiativeChanged.value += "{" +  NoCase + "};";
	}
}

function GoalsChange() {
	var temp = event.srcElement.id.split("_");
	var NoCase = temp[1];
	
	if(document.forms[0].GoalsChanged.value.indexOf("{" + NoCase + "}") < 0) {
		document.forms[0].GoalsChanged.value += "{" +  NoCase + "};";
	}
}

function CheckUncheck() {
	var temp = event.srcElement.id.split("_");
	
	if(temp[0] == "A" && document.getElementById("A" + "_" + temp[1]).checked == true)  {
		document.getElementById("B" + "_" + temp[1]).checked = false;
	} else if(temp[0] == "B" && document.getElementById("B" + "_" + temp[1]).checked == true)  {
		document.getElementById("A" + "_" + temp[1]).checked =  false;
	}
}

