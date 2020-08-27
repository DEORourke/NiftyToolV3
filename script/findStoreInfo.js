var sharepointURL = 'http://teams.company.com/sites/site/Files/';

var storeData;
var keys = {};
var lookupHistory = new Array();
var histPos
var techName = ""
var ups;
	
$(document).ready(function () {
	updateDataTimer();
	resizeElements();
	getJulianDate();
	retrieveCookies();
	hideSections();
	apf();
	$("#storeNum").focus();

	$(window).resize(function() {resizeElements()});
	
	$(document).on("keydown", function (b) {
		keys[b.which] = true;
		eventFire();
	});
	
	$(document).on("keyup", function (b) { delete keys[b.which]; });
	
	$("#remote").change(function() {launchRemoteOptions()});
	movePopOut($(".popOutHeader"), $(".popOut"))
	
	$("#logo").click( function() { pass() });
});

function pass() {
	if (keys[113]) { ups = true } 
}

function hideSections() {
// called on page load. Hides the different divs on page. Also gets the width of the phone and email divs and moves them off screen.

	$(".infoTable").hide();
	$(".remoteOptions").hide();
	$(".phoneList").hide(2, function() {
		var pWidth = "-=" + (parseInt($(".phoneList").css("width").replace(/px/,"")) + 25) + "px";	// 25px added to the transform to compensate for the scroll bar.
		$(".phoneList").animate({right: pWidth },2,function() {
			$(".phoneList").show();
		});
	});
	$(".emailTemplates").hide(2, function() {
		var eWidth = "-=" + (parseInt($(".emailTemplates").css("width").replace(/px/,"")) + 15) + "px"; // 15px added to the transform to compensate for the shadow.
		$(".emailTemplates").animate({left: eWidth },2,function() {
			$(".emailTemplates").show();
		});	
	});
	$("#history").hide();
	$(".labLinks").hide();
	$(".quickLinks").hide();
	$(".popOut").hide();
	$(".immediateConnectOptions").hide();
};


function eventFire() {
// called by the document.ready script listening for key press. Evaluates depressed keys and calls functions if conditions are met.

	var focusType = $(':focus').attr('type');
	var template = $('#popOutTitle').text();
	var numberKeys = [49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105];


	if (keys[13]) {
		//Other things to check for if enter is pressed
		if (($("#storeNum").is(":focus") || $("#remote").is(":focus")) && $('#storeNum').val().trim().length > 0) {
			grepData($('#storeNum').val().trim());
			keys = {};	//clears array containing keys pressed to avoid double error messages.
		} else if ( (template === "List stores by...") && (focusType === "text" || focusType === "radio" || focusType === "select") ) {
			activateListStores();
			keys = {};
		} else if ( (template !== " ") && (focusType === "text" || focusType === "radio" || focusType === "checkbox") ) {
			sendEmail();
			keys = {};
		};
		
	} else if (keys[18]) {
		//Other things to cehck for if Alt is pressed
		if (keys[81]) { togglePopInSections("emailTemplates");
		} else if (keys[80]) { togglePopInSections("phoneList");
		} else if (keys[76]) { toggleSections("labLinks");
		} else if (keys[79]) { toggleSections("docLinks");
		} else if (keys[87]) { toggleSections("webLinks");
		} else if (keys[85]) { $("#updateButton").click();
		} else if (keys[89]) { writeListStores();
		} else if (keys[77]) { writeMassConnect();
		} else if (keys[82] && template !== " ") { sendEmail() };
		
	} else if (keys[27]) { clearMid();
	} else if (keys[119]) { $("#storeNum").focus();	//F8 key moves cursor to store number field
	} else if (keys[38] && $("#storeNum").is(":focus")) {	//cycling through history
			navigateHistory(0);
	} else if (keys[40] && $("#storeNum").is(":focus")) {	//cycling through history the other direction
			navigateHistory(1);
	};

};


function updateData(lastChecked) {
// called by the updateDataTimer function. Gets headers of the three CSV files from SharePoint, then compares the Last-Modified value to the last time it was checked, and updates if needed.
// Accepts one argument: lastChecked - An int containing the timestamp of the last time this function was run.
	
	var storeDataURL = sharepointURL + "NiftyStoreInfo.csv";
	var partsListURL = sharepointURL + "NiftyPartsInfo.csv";
	var phoneListURL = sharepointURL + "NiftyPhoneInfo.csv";
	var labPCListURL = sharepointURL + "NiftyLabPCInfo.csv";
	var quickLinkURL = sharepointURL + "NiftyQLinkInfo.csv";

	console.log("Checking for updated information at " + Date(Date.now()));
	
	function checkAndUpdateData(URL, nameOfFunction, listName){
		$.get(URL,function(data, statusMsg, request){
			let infoModified = Date.parse(request.getResponseHeader("Last-Modified"))
			if ((storeData == undefined) || (infoModified > lastChecked)) {
				eval(nameOfFunction + "(data)");
			};
		}).fail(function(data) {
			alert("Unable to load " + listName + ".\nCall to get from Sharepoint failed with error: " + data.statusText);
		});
	};

	checkAndUpdateData(partsListURL, "loadPartsList", "parts list");
	checkAndUpdateData(phoneListURL, "drawPhoneList", "phone list");
	checkAndUpdateData(labPCListURL, "drawLabRemote", "lab PC list");
	checkAndUpdateData(quickLinkURL, "drawQuickLinks", "quick links");
	checkAndUpdateData(storeDataURL, "loadStoreFile", "store info");	
}


function updateDataTimer() {
// Called on document.ready. Calls the updateData function once, then again every 7 minutes.

	var lastChecked = 0
	updateData(lastChecked);
	lastChecked = Date.now();
	
	window.setInterval(function() {
		lastChecked = updateData(lastChecked)
		lastChecked = Date.now();
	}, 420000);
}


function loadStoreFile(data) {	
// called by the updateData function. Parses the NiftyStoreInfo.csv file and updates the storeData global variable.
// Accepts one arguement: data - contents of the NiftyStoreInfo.csv file.

	Papa.parse(data, {
		delimiter: ',',
		quoteChar: '"',
		header: true,
		complete: function(results) {
			console.log("Store data loaded and parsed:", results);
			storeData = results.data;
		}
	});
};


function resizeElements() {
	// called on page load and by the writePartsOrder function. Resizes the phone list and popout elements for display on laptops.
	
	let windowSpace = window.innerHeight - $(".header").height() - parseInt($(".header").css("padding-top")) - parseInt($(".header").css("padding-bottom"));
	let phoneListPadding = parseInt($(".phoneList").css("padding-top")) + parseInt($(".phoneList").css("padding-bottom"));
	let phoneListMaxSize = parseInt($(".phoneList").css("max-height")) + phoneListPadding;
	let emailListPadding = parseInt($(".emailTemplates").css("padding-top")) + parseInt($(".emailTemplates").css("padding-bottom"));
	let emailListMaxSize =  parseInt($(".emailTemplates").css("max-height")) + emailListPadding;
	let popOutPadding = 1
	let popOutMaxSize = parseInt($(".popOut").css("max-height"));
		
	if (windowSpace < phoneListMaxSize) {
		let newPhoneListSize = windowSpace - phoneListPadding - parseInt($(".phoneList").css("border-radius"))
		$(".phoneList").height(windowSpace - phoneListPadding - 10);
		$(".phoneTable").height(windowSpace - phoneListPadding - 10 - 35);
	} else if (windowSpace > phoneListMaxSize && windowSpace < phoneListMaxSize +5 ) {
		$(".phoneList").css("height","");
		$(".phoneTable").css("height","");
	}
		
	if (windowSpace < emailListMaxSize) {
		$(".emailTemplates").height(windowSpace - emailListPadding - 10);
	} else if (windowSpace > emailListMaxSize && windowSpace < emailListMaxSize +5 ) {
		$(".emailTemplates").css("height","");
	}
	
	if (windowSpace < popOutMaxSize ) {
		$(".popOut").height(windowSpace - popOutPadding - 10);
	
	} else if (windowSpace > popOutMaxSize && windowSpace < popOutMaxSize +5 ) {
		$(".popOut").css("height","");
	}
};


function drawPhoneList(data) {
	Papa.parse(data, {
		delimiter: ',',
		quoteChar: '"',
		header: true,
		complete: function(results) {
			phoneList = results.data
			var Header = "";
			var blankCell = "<td>&nbsp;</td>"
			var RetailSupportCounter = 0;
			var CompanyContactsCoutner = 0;
			var VentorContactCounter = 0;
			console.log("Phone list loaded and parsed:", results);
			
			$("#phoneHeader").html("\
				<tr>\
					<th class='col1'>Contact</th>\
					<th class='col2'>IM</th>\
					<th class='col3H'>Office</th>\
					<th class='col3H'>Mobile</th>\
					<th class='col4'>Email</th>\
				</tr>\
			");
			
			$('.phoneTable table').text('')
			
//Is there a better way than .each to loop through these?				
			$.each(phoneList,function(i, n) {
				let nameCell
				let imCell
				let officeCell
				let mobileCell
				let emailCell

				if ( n["Division"] != Header) {	// creating section headers
					Header = n["Division"];
					$("#" + n["Category"] + "Table").append("<tr><th colspan='4'>" + Header + "</th></tr>");
				};

				if ( !n["Notes"] == "") {
					nameCell = "<td class='col1'>" + n["Contact"] + "<img class='noteIndicator' src='./graphic/note2.png' /><span class='noteText'>" + n["Notes"] + "</span></td>";
				} else {
					nameCell = "<td class='col1'>" + n["Contact"] + "</td>";
				};

				//fill in table information with the correct class if the value isn't blank
				if ( !n["IM"] == "") {
					imCell = "<td class='col2'><a href='sip:" + n['IM'] + "'><img src='./graphic/chatico2.png' /></a></td>";
				} else { imCell = blankCell; };

				if ( !n["Office"] == "") {
					officeCell ="<td class='col3 link' id='O" + i + "' onclick='dialOut(this.id)'>" + n["Office"] + "</td>";
				} else { officeCell = blankCell; };

				if ( !n["Mobile"] == "") {
					mobileCell = "<td class='col3 link' id='M" + i + "' onclick='dialOut(this.id)'>" + n["Mobile"] + "</td>";
				} else { mobileCell = blankCell; };

				if ( !n["Email"] == "") {
					emailCell = "<td class='col4 link'><a href='mailto:" + n['Email'] + "' target='_top'>" + n['Email'] +"</a></td>";
				} else { emailCell = blankCell; };

				let tableRowName = "contactListLine" + i; //Create a new row in the appropriate table for the contact and give it a name.
				
				$("#" + n["Category"] + "Table").append("<tr id='" + tableRowName + "'>" + nameCell + imCell + officeCell + mobileCell + emailCell + "</tr>");
					
				switch (n["Category"]) { // setting off-color for every other row of each table
					case "RetailSupport":
						RetailSupportCounter++;
						if ( RetailSupportCounter % 2 == 0) { $("#" + tableRowName).addClass("oddRow") };
						break;
					case "CompanyContacts":
						CompanyContactsCoutner++;
						if ( CompanyContactsCoutner % 2 == 0) { $("#" + tableRowName).addClass("oddRow") };
						break;
					case "vendorContacts":
						VentorContactCounter++;
						if ( VentorContactCounter % 2 == 0) { $("#" + tableRowName).addClass("oddRow") };
						break;	
				};

			});
		changePhoneList("retailSupportListButton");
				
		$(".noteIndicator").hover(function() {
			let hoveredObject = "#" + $(this).parents('tr').attr("id");
			//let windowObject = window.event
			let Xpos = window.event.clientX
			let Ypos = window.event.clientY
//can use toggleClass()?				
			$(hoveredObject).find(".noteText").css("visibility", "visible");
			$(hoveredObject).find(".noteText").offset({top: Ypos - 10, left: Xpos + 15});
		}, function() {
			let hoveredObject = "#" + $(this).parents('tr').attr("id");
			$(hoveredObject).find(".noteText").css("visibility", "hidden");	
		});			
		}
	});
	
	phoneList = undefined;
};


function changePhoneList(buttonID) {
//use hasclass and toggle class for efficiency sake?	
	var listID = "#" + buttonID.replace("ListButton", "Container")
	$(".phoneTable").hide();
	$(".phoneListButton").removeClass("phoneButtonClicked");
	$(listID).show();
	$("#" + buttonID).addClass("phoneButtonClicked");
};


function grepData(search) {	
// called on demand. Searches the store and orical fields of the store data CSV for the value passed to it, stores it, then passes the info to the fillTable function.
	//Needs to be rewritted so other functions call this one and it returns the row to them
	$("#history").hide();
	tableReset();

	$("#storeNum").val("");
	var data = $.grep(storeData, function(line) {	// When that text's in the string, Ma', grep it like it's hot.
		return (line.Store == search || line.Orical == search);
	});
	try {
		if(isNaN(search) || search == "") throw "Enter a numeric store number.";
		if(data[0] === undefined) throw search + " is not a valid store number.";
	} catch(erMsg) {
		document.title = "Nifty Tool v3";
		alert( erMsg );
		$("#storeNum").focus();
		return;
	};	
	modifyHistory(search);
	fillTable(data[0]);
	if(keys[16]){ dialOut("phone") };	//dials out if shift is held
};


function fillTable(info) {	// called by the grepData function. Fills in the cells of the infoTable div with store information and stores information needed by other functions in a hidden div on page to avoid the use of global variables.
	document.title = info.Store + " - Nifty Tool v3";

	$("#storeNumber").text(info.Store);
	$("#storeOrical").text(info.Orical);
	if(info.Retailer != "") {
		if(info.Retailer.length > 21) {
			$("#retailer").text(" - " + info.Retailer.substring(0, 19) + "..");
		} else {
			$("#retailer").text(" - " + info.Retailer);
		}
	} else {
		$("#retailer").text("");
	};
	if(info.Phone != "" && info.Phone != "None") {	// cell only get the link class if it has information
		$("#phone").html('<span class="lnk">' + info.Phone + '</span>');
	} else {
		$("#phone").html("");
	};
	$("#router").text(info.Router);
	$("#MFS1").text(info.MFS1);
	$("#switch_2").text(info.Sw2);
	$("#cpuser").text(info.CPuser);
	
	if (info.ClosedDate != "") {
		$("#hours").text("Closed " + info.ClosedDate);
	} else {
		if (info.NumLanes == "") {
			$("#versionLanes").text("Stand Beside Store");
		} else {
			$("#numLanes").text(info.NumLanes);
			$("#versionLanes").text(info.iss45v);
			$("#makeLanes").text(info.POSmake + " lanes");
		};
	
		if(info.Hours != "" && info.SundayHours != "") {	// either fills in both hours cells, or neither if some info is missing.
			$("#hours").text("M-S : " + info.Hours + " - " + info.TimeZone);
			$("#sunHours").text("Sun : " + info.SundayHours);
		} else {
			$("#hours").text("");
			$("#sunHours").html("&nbsp");
		};
	};
	
	$("#access_point").text(info.AP);
	$("#cpcompany").text(info.CPcompany);
	
	if (info.Broadband == "BYO") {
		if ((info.DS1circuitID.search("DS1ITDT-")) == 0) {
			$("#broadband").text("BYO - Company Provides Backup");
		} else {
			$("#broadband").text("BYO Only");
		};
	} else {
		$("#broadband").text(info.Broadband);
	};

	$("#storeAddress").text(info.Street);
	if (info.City != ""){ $("#storeCity").text(info.City + ", ") };
	$("#storeState").text(info.State);
	$("#storeZip").text(info.Zip);
	$("#apModel").text(info.APModel);
	$("#dc").text(info.DC);
	$("#region").text(info.Region);
	if(info.Retailer === "Corporate") {	// avoids filling in data for equipment licensee stores don't have. It's on the Rollout sheet, so it's in the table.

		$("#MIO_PC").html('<div class="lnk">' + info.MIO_PC + '</div>');
		$("#safe").text(info.Safe);
		$("#meat_scale").html('<div class="lnk">' + info.Meat + '</div>');
		$("#produce_scale").html('<div class="lnk">' + info.Produce + '</div>');
		if(info.InTouch) {
			$("#time_clock").html('<div class="lnk">' + info.Time + '</div>');
		} else {	
			$("#time_clock").text(info.Time);
		};
	} else {
		$("#MIO_PC").text("");
		$("#time_clock").text("");
		$("#safe").text("");
		$("#meat_scale").text("");
		$("#produce_scale").text("");
		
	};
	$("#cppass").text(info.CPpass);
	$("#routermodel").text(info.RouterModel);

	if (!info.HDPilot == "") {	// Is the store part of the new helpdesk pilot?
		let added = new Date(info.HDPilot);
		let now = new Date()
		if (added < now) {
			$("#HDPilot").text("L1 helpdesk pilot store");
		} else {
			$("#HDPilot").text("");
		};
	} else {		
		$("#HDPilot").text("");
	};
	
	$(".infoTable").slideDown(200, function() {
		if ($("#remote").is(":checked")) {	// option to immediately connect to the store server.
			connectOption = $("input[name='immediateConnectOptions']:checked").val();
			remoteConnect(info[connectOption], connectOption);
			
		} else {
			$("#connectOption0").prop("checked", true);	// set default connection back to dashboard if immediate connect is not checked.
		};
	});
	
		//stored data fields
	$("#storeCPOld").text(info.oldCPcompany);
	$("#urlException").text(info.URL_Exception);
	$("#hasIntouch").text(info.InTouch);

};


function tableReset() {	// called by the grepData function. Hides and clears (most values in) the store info table and as well as the contents of the stored values elements.
	$(".infoTable").hide();
	$(".storeValues span").each(function(i, n) { $(n).empty() });
	$(".infoTableMain span").each(function(i, n) { $(n).empty() });
}


function pingDevice(address) {	// called on demand. Reads the text of the ID the button corresponds to and opens a shell to ping the address using the desired option.
	var cmdShell = new ActiveXObject("wscript.Shell");
	var ipaddress = $(address).text();
	var opt = $("input[name='pingOption']:checked").val();
	var store = $("#storeNumber").text();
	var device = address.substring(1).replace(/_/, " ");
	var time = formatTime(2);
	var title = "Pinging Company " + store + " " + device;
	
	if ($("input[name='pingSize']:checked").val() == 1) {
		var size = 1024;
	} else {
		var size = 32;
	};
	
	if (ipaddress == "") {
		return;
	} else {
		switch (opt) {
			case "0":
				cmdShell.Run('cmd /c title ' + title + ' && ping -l ' + size + ' ' + ipaddress + ' && pause');
				break;
			case "1":
				cmdShell.Run('cmd /c title ' + title + ' && ping -l ' + size + ' -t ' + ipaddress);
				break;
			case "2":
				
				var path = '%userprofile%\\documents\\'
				var file = 'ping_Company' + store + device + '_' + time + '.txt';
				cmdShell.Run('cmd /c title =' + title + ' && echo NiftyTool will now attempt to save ' + file + ' to the && echo ' + path + ' folder, but will need permission. && echo. && runas /user:domain\\%username% "\\\\server\\path\\to\\LogPingWithTimestamp.bat "' + ipaddress + '" ' + size + ' ' + path + file);
		//	Would really liked to have found a way to run this without referencing an external batch file. Maybe in a future update. The issue is that I'd need to have nested quotes 4 deep for both this function and the contents of the batch file.
				break;
		};
	};
};


function pingAllDevices() {	// called on demand. Compiles a list of devices with ID addresses in the store information table and passes them to a script that pings all the devices.
	var cmdShell = new ActiveXObject("wscript.Shell");
	var deviceList = "";
	var fieldList;
	var title
	
	title = "Pinging devices at store " + $("#storeNumber").text();
	fieldList = ["router", "MFS1", "switch_2", "access_point", "safe", "meat_scale", "MIO_PC", "time_clock", "produce_scale"]

	fieldList.forEach( function(i, n) {

		if ( $("#" + i).text() != "" ) {
			deviceList = deviceList + $("#" + i).text() + "-" + i + " " ;
		};
	});
	
	cmdShell.Run('cmd /c title ' + title + ' && cscript "X:\\path\\NP3\\script\\PingStoreDevices.vbs" ' + deviceList + ' && pause');
}


function launchVNC(address, devType) {
// Launches VNC to connect to scales or intouch model time clocks.
// Accepts two arguments: address - the IP address of the device to connect to
//				devType - device type. This is the text of the element ID to this function from the onclick event.

	if(devType == "time_clock" && !$("#hasIntouch").text()){return};	// keeps the function from trying to connect to 4500 model time clocks

	var cmdShell = new ActiveXObject("wscript.Shell");
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	
	var b32Path = "C:\\Program Files (x86)\\RealVNC\\VNC Viewer\\vncviewer.exe"
	var b64Path = "C:\\Program Files\\RealVNC\\VNC Viewer\\vncviewer.exe"

	if(fso.FileExists(b64Path)) { cmdShell.run('"' + b64Path + '"' + address); }
	else if(fso.FileExists(b32Path)) {	cmdShell.run('"' + b32Path + '"' + address); }
	else {
		alert("VNC Viewer is not installed in the default location. \nUnable to launch VNC Viewer. Woops.");
	};
}


function getJulianDate() {
// Returns yesterdays Julian calendar date.

	var now = new Date();
	var Y = now.getFullYear();
	var M = now.getMonth();
	var D = now.getDate();
	var E = 0;	//exceptions
	
	var mon30 = [1, 3, 5, 7, 8, 10];
	var mon31 = [2, 4, 6, 9, 11];

	if (mon30.indexOf(M) > -1) {	// Adding extra days for past 31 day months
		E = mon30.indexOf(M) + 1;
	} else if (mon31.indexOf(M) > -1) {
		E = mon31.indexOf(M) + 1;
	};
	if (M >= 2) {	
		E -= 2;	// Subtracting two days for February
		if (Y % 4 === 0) {	// Accounting for leap year
			E += 1; 
		};
	};

	M *= 30;

	if ( M == 0 && D == 1 ) {	// Sets the correct date for yesterday if today is Jan 1
		if (Y % 4 === 1) { E = 1 };
		var Yesterday = 365 + E;
	} else {
		var Yesterday = M + D + E - 1;
	}
	
	$("#YJD").text(Yesterday);
};


function formatTime(opt) {
// called by pingDevice function. Returns the current time.
	var now = new Date();
	var time;
	var Y = now.getFullYear();
	var M = now.getMonth() + 1;
	var D = now.getDate();
	var dy = now.getDay();
	var h = now.getHours();
	var m = now.getMinutes();
	var s = now.getSeconds();
	
	if (D < 10) { D = "0" + D };
	if (M < 10) { M = "0" + M };
	
	switch (opt) {
		case 1:
			time = h + "" + m + "" + s;
			break;
		case 2:
			time = M + "" + D + "-" + h + "" + m + "" + s;
			break;
		case 3:
			time = M + "/" + D + "/" + Y;
			break;
		case 4:
			time = M + "/" + D;
			break;
	};
	return time;
};


function remoteConnect(address, computer) {
// New version of remote connect feature 
// Accepts two arguments (ip address and device to connect to)
	var app = $("input[name='connectOption']:checked").val();

	switch (app) {
		case "0":
			app = "/main.html";
			break;
		case "1":
			app = "/remctrl.html";
			break;
		case "2":
			app = "/filexfer.html";
			break;
		case "3":
			if (ups) {
				launchCMD(address, computer);			
				return;
			} else {
				app = "/telnet.html";
			};
			break;
		case "4":
			app = "/evtlogs.html";
			break;
	};
	if (computer == "MIO_PC") {
		var port = ":4600";
		if (ups) { app = "/login:creds" + app };
	} else {
		var port = ":2100";
		if (ups) { app = "/login:creds" + app };
	};

	if (address == "") {
		return;
	} else {
		window.open("https://" + address + port + app);
	};
}


function drawLabRemote(data) {
// called by the updateData function. Parses the NiftyLabInfo.csv file and creates links for each entry in the LabLinks table. Also attaches an event listener to the labLinks class
// Accepts one argument: data - contents of the NiftyLabInfo.csv file.

	Papa.parse(data, {
		delimiter: ',',
		quoteChar: '"',
		header: true,
		complete: function(results) {
			var labList = results.data.slice(0, results.data.length - 1);
			var labLinkDiv = $("#labLinkAddresses");
			console.log("Lab PC list loaded and parsed:", results);
			labLinkDiv.text("")
			$.each(labList,function(i, n) {
				labLinkDiv.append('\
					<a class="labLink" id="lablink' + i + '" target="_blank" href="https://' + n["IP Address"] + ':' + n["RA Port"] + '/"><span>' + n["Label"] + '</span></a>\
				');
			});
		}
	});
	$(".labLink").off("click");
	$(".labLink").click( function(event) { labRemote(this.id, event) });
	labList = undefined;
}


function labRemote(id, trigger) {
// called by the event listener set up in the drawLabRemote function. If flag is set, opens hyperlink with quick-login info. Also calls sendLabCallout function from emailtemplates if the checkbox is set.
// Accepts two arguments: ID - the ID of the lab link being clicked.
//				trigger - the event triggering this function to be called. the click.

	if (ups) {
		trigger.preventDefault();
		
		var labLink = $("#" + id);
		var addr = labLink.attr('href');
		if (addr.substr(addr.lastIndexOf(":")) == ":4600/") {
			addr += "login:creds";
		} else {
			addr += "login:creds";
		};
		
		window.open(addr);
	};
	
	if ($("#labLinksCheckbox").is(":checked")) { sendLabCallout(id) };
};


function drawQuickLinks(data) {
// called by the updateData function. Parses the NiftyQLinksInfo.csv file and creates hyperlinks for each item.
// Accepts one argument: data - contents of the NiftyQLinksInfo.csv file.

	Papa.parse(data, {
		delimiter: ',',
		quoteChar: '"',
		header: true,
		complete: function(results) {
			console.log("Quicklink list loaded and parsed:", results);
			
			var linkList = results.data.slice(0, results.data.length - 1);
			var quickLinksContainer = $("#quickLinksContainer")
			var quickLinkRows = []
			var rows = []
			
 			$.each(linkList, function(i, n){	// collect unique row numbers in an array
				if (rows.indexOf(n["Row"]) == -1) { rows.push(n["Row"]) };
			});
			
			rows.sort()
			
			quickLinksContainer.text("");
			$.each(rows, function(i, n) { 	// create a div in the parent container for each row. Select the row and store it in an array
				quickLinksContainer.append('<div class="row" id="quickLinkRow' + n + '"></div>');
				quickLinkRows[n] = $("#quickLinkRow" + n);
			});		
			
			$.each(linkList, function(i, n) {	// put each item in the appropriate row
				let target = "";
				if ((n["Address"].substr(0,2) != "\\\\") && (n["Address"].substr(1,1) != ":")) {
					target = 'target="_blank" ';
				};	// If the address is a server path, the link won't target a new window
				
				quickLinkRows[n["Row"]].append('\
					<div class="cell">\
						<a ' + target + 'href="' + n["Address"] + '">\
							<img src="graphic/' + n["Icon"] + '">\
							<p>' + n["Label"] + '</p>\
						</a>\
					</div>\
				');
			});
		}
	});
	linkList = undefined;
}


function showWiringGraphic(field) {
// called by onclick event. Reads the text of the field the clicked button corresponds to and launches a new browser tab with the appropriate graphic of POS or network wiring layout.
// Accepts one argument: field - The ID of the field containing the name of the desired image.

	var string = $(field).text();
	var path = "\graphic\\";
	if (string.search(/HP lanes/i) > -1) { window.open(path + "HP.png");
	} else if (string.search(/NCR lanes/i) > -1) { window.open(path + "NCR.png");
	} else if (string.search(/Cisco 4321/i) > -1) { window.open(path + "Cisco4321.vsd");
	} else if (string.search(/Cisco 1841/i) > -1) { window.open(path + "Cisco1841.png");
	} else if (string.search(/Cisco 1941/i) > -1) {window.open(path + "Cisco1941.png");
	};
};


function dialOut(field) {
// Called by on click event. Gets the text of the ID that called it and launches a link to dial the number.

	var number = $("#" + field).text();

	var prefix = $("input[name='location']:checked").val();
	var address = $("#address2").text();
	try {
		if (isNaN(prefix)) throw "Select a location.";	// Offices have two different dial out prefixes. This makes sure the correct one gets applied.
		if (number === undefined) throw "No number found.";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	if (address.search(/Aruba/i) > -1) {	// adds the appropriate country code for calling Aruba if information for one of those stores is in the table.
		var cCode = "011";
	} else {
		var cCode = "1";
	};
	window.open("tel:" + prefix + cCode + "-" + number, "_self");
};


function toggleSections(id) {	// called on demand. Receives the id of the button that called it to show/hide the associated div and change the button text.

	$("." + id).slideToggle(200);
	id = "#" + id
	var field = $(id).text();
	var string
	if(field.search(/show/i) == -1) {
		string = $(id).html().replace("Hide", "Show");
	} else {
		string = $(id).html().replace("Show", "Hide");
	};
	$(id).html(string);
};


function togglePopInSections(id) {	// called on demand. Receives the id of the button that called it to animate the associated div and change the button text.
	var cls = "." + id
	id = "#" + id
	var field = $(id).text();
	var width = $(cls).css("width");
	var string
	var tWidth = parseInt(width.match(/\d+/)) + 15;
	
	width = width.replace(width.match(/\d+/),tWidth);
	if (field.search(/show/i) == -1) {
		width = "-=" + width;
		string = $(id).html().replace("Hide", "Show");
	} else {
		width = "+=" + width;
		string = $(id).html().replace("Show", "Hide");
	};
	if (id == "#phoneList") {
		$(cls).animate({right: width},200);
	} else {
		$(cls).animate({left: width},200);
	};
	$(id).html(string);
	
};


function weeklyAd() {
// called by clicking the Store Weekly Ad button. Reads the value from the storeNumber ID and launches a new tab to the store's weekly ad page.

	window.open("https://Company.com/weeklyad?store_code=" + $("#storeNumber").text());
};


function storeWebPage() {
// called by clicking the Store Web Page button. Checks to see if there is a URL exception for the currently loaded store, then constructs and clicks the URL for the store's web page.

	var storeNumber = $("#storeNumber").text();
	var city = $("#storeCity").text();
	var zipcode = $("#storeZip").text();
	var urlException = $("#urlException").text();
	//var urlException = $("#storeURL").text();
	var URL = "https://Company.com/stores/";

	try{
		if(urlException == "SBS" | urlException == "Franchise") throw "The URL for this stand-beside store is unknown."
	}catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	city = city.replace(/\./g, " ").replace(/  /g," ").replace(/ /g, "-");
	if (urlException == ""){
		URL = URL + city + "-" + zipcode + "-" + storeNumber + "/";
	} else {
		URL = URL + urlException;
	};

window.open(URL)
}


function modifyHistory(addition) {	// called by the grepData function. Adds the search store to the beginning of an array and removes the last item if the length is now greater than 10.
	lookupHistory.unshift(addition);
	if (lookupHistory.length > 10) { lookupHistory.pop() };
	writeHistory();
	document.cookie = "lookupHistory = " + lookupHistory + ";";
	histPos = null;	// resets the position for the navigateHistory function.
};


function writeHistory() {	// called by modifyHistory and retrieveCookies functions. draws links to re-search for the stores in the history element.
	$(".historyBanner").text("");
	$.each(lookupHistory, function(i, n) {
		$(".historyBanner").append("<a class='historyLink' onclick='grepData(this.text)'>" + n + "</a>&nbsp;&nbsp;")
	});
}


function recallHistory() {	// called on demand. Toggles visibility of the element containing the store search history.
	$("#history").toggle();
};


function navigateHistory(direction) {	// called on demand. Fills the storeNum field with the next or previous item in the lookupHistory array.

	if ( lookupHistory.length == 0 ) { return; };	// You can't navigate through what's not there.
	var value = $("#storeNum").val();

	switch (direction) {
		case 0:	// Going backwards.
			if ( !$.isNumeric(value) || !lookupHistory.indexOf(value) == -1 ) {
				$("#storeNum").val(lookupHistory[0]);
				histPos = 0;
			} else if ( histPos == lookupHistory.length - 1 ) {	// Clears the input field if at the end of the array
				$("#storeNum").val("");
			} else {
				histPos++;
				$("#storeNum").val(lookupHistory[histPos]);
			};
		break;
		
		case 1:	// Going forward.
			if ( !$.isNumeric(value) || !lookupHistory.indexOf(value) == -1 ) {
				$("#storeNum").val($(lookupHistory).get(-1));
				histPos = lookupHistory.length - 1;
			} else if ( histPos == 0 ){	// Clears the input field if at the begining of the array
				$("#storeNum").val("");
				histPos = null;
			} else {
				histPos--;
				$("#storeNum").val(lookupHistory[histPos]);
			};
		break;
	};
};


function retrieveCookies() {	// called on document load. Reads cookies and extracts the lookupHistory information.
	var posHistory = -1
	var posTech = -1
	var posLocation = -1
	var allTehCookies = document.cookie.split("; ");

	$.each(allTehCookies, function(i,n) {
		if (!n.search(/^lookupHistory/i)) { return posHistory = i }
		if (!n.search(/^techName/i)) { return posTech = i }
		if (!n.search(/^location/i)) { return posLocation = i }
	});
	
	if (posHistory >= 0) { 
		lookupHistory = $(allTehCookies)[posHistory].replace(/lookupHistory=/,"").split(",");
		writeHistory();
	};
	if (posTech >= 0) {
		techName = $(allTehCookies)[posTech].replace(/techName=/,"")
	};
	if (posLocation >= 0) {
		//something something set location
	};
}


function launchCMD(ipaddress, computer) {	// called by the remoteConnect function. Launches a window resized for RemotelyAnywhere's host command prompt.
	if (computer == "mio") {
		window.open("https://" + ipaddress + ":4600/login:creds/telnet.html", "", "width=810,height=550");
	} else {
		window.open("https://" + ipaddress + ":2100/login:creds/telnet.html", "", "width=810,height=550");
	};
};


function apf() {
	var day = new Date;
	var d = day.getDate();
	var m = day.getMonth();
	if (d == 1 && m == 3) { $('img').css("transform", "rotate(180deg)"); };
}

function hideStoreInfo() {	// called on demand. Hides the infoTable element and sets the storeNumber field value to the currently viewed store.
	$(".infoTable").hide();
	$("#storeNum").val($("#storeNumber").text());
	$("#getInfo").focus();
};


function writeListStores() {	// called on demand. Draws the input fields for the doListStores function in the popout element and unhides it, or clears it if it's already up.
	if ( $(".popOut").is(":visible") && $("#popOutTitle").text() == "List stores by..." ) {
		clearEmail();
		return;
	};
	
	var loadedCompany = $("#cpcompany").text();
	var loadedState = $("#storeState").text();
	var loadedRegion = $("#region").text();
	var loadedDC = $("#dc").text();
	
	var allStates = allStateSelector(loadedState);
	var allRegions = allRegionSelector(loadedRegion);
	var allDCs = allDCSelector(loadedDC);
	
	$("#popOutTitle").text("List stores by...");
	$(".popOutBody").html("<div class='listStoresOptions'></div><div class='listStoresContainer'></div>");

	$(".listStoresOptions").html("\
		<input type='radio' name='lookupOption' id='lookupByState' value='0'><label for='lookupByState'> State</label>&nbsp; \
		<input type='radio' name='lookupOption' id='lookupByCompany' value='1'><label for='lookupByCompany'> Company</label>&nbsp; \
 		<input type='radio' name='lookupOption' id='lookupByRegion' value='2'><label for='lookupByRegion'> Region</label>&nbsp; \
 		<input type='radio' name='lookupOption' id='lookupByDC' value='3'><label for='lookupByDC'> DC</label> : &nbsp; \
		\
		<select class='listByOption' id='lookupByStateField'>" + allStates + "</select>\
		<input class='listByOption' id='lookupByCompanyField'  maxlength=6 size='6' value='" + loadedCompany + "' />\
		<select class='listByOption' id='lookupByRegionField'>" + allRegions + "</select>\
		<select class='listByOption' id='lookupByDCField'>" + allDCs + "</select>\
		\
		<span class='listStoresButton'><button class='listStoresButton' id='lookupButton'>List</button></span>\
	");
	
	$(".listStoresContainer").html("\
		<b id='topLine'></b>\
		<span class='listStoresButton'>\
			<button onclick='exportListStores()'>Export</button>&nbsp;\
			<button onclick='listStoresMassConnect()'>Connect</button>\
		</span>\
		\
		<br /><br />\
		\
		<table id='listStoresHeader'>\
			<td  class='storesColNarrow' onclick='sortTable(1)'>\
				Store \
				<span class='sortTable' id='sortTable1'>\
					<img src='./graphic/arrow3.png'>\
				</span>\
			</td>\
			<td class='storesColWide'>\
				Server\
			</td>\
			<td class='storesColWide' onclick='sortTable(3)'>\
				City \
				<span class='sortTable' id='sortTable3'>\
					<img src='./graphic/arrow3.png'>\
				</span>\
			</td>\
			<td class='storesColNarrow' onclick='sortTable(4)'>\
				State \
				<span class='sortTable' id='sortTable4'>\
					<img src='./graphic/arrow3.png'>\
				</span>\
			</td>\
			<td class='storesColWide'>\
				Phone \
			</td>\
			<td class='storesColNarrow'>\
				<input type='checkbox' id='masterRCCheckbox' /> \
				Remote\
			</td>\
		</table>\
		\
		<div class='listStores'>\
			<table id='listStores'>\
			</table>\
		</div>\
	");
	
	
	$("input[name='lookupOption']").change(function() {
		$(".listByOption").hide();
		switch ($("input[name='lookupOption']:checked").val()) {
			case "0":
				$("#lookupByStateField").show();
				break;
			case "1":
				$("#lookupByCompanyField").show();
				break;
			case "2":
				$("#lookupByRegionField").show();
				break;
			case "3":
				$("#lookupByDCField").show();
				break;		
		};
	});
	
	$("#lookupButton").click( function(){ doListStores(); });
	$(".listStoresContainer").hide()
	$(".listByOption").hide();
	$(".popOut").show();
};


function doListStores(search) {	// called by activateListStores function. Determines the field to search by, then searches the storeData array for matches before drawing them to a table.
	
	var lookupOption = $("input[name='lookupOption']:checked").val();
	var search;
	var field
	var topLine;
	
	switch (lookupOption) {
		case "0":
			search = $("#lookupByStateField").val();
			field = "State";
			topLine = " stores in " + search;
			break;
		case "1":
			search = $("#lookupByCompanyField").val();
			field = "CPcompany";
			topLine = " Stores in Company " + search;
			break;
		case "2":
			search = $("#lookupByRegionField").val();
			field = "Region";
			topLine = " Stores in Region " + search;
			break;
		case "3":
			search = $("#lookupByDCField").val();
			field = "DC";
			topLine = " Stores serviced by DC " + search;	
			break;
	};

	var data = $.grep(storeData, function(line) {
		return (line[field] === search);
		});
	$("#topLine").text(" " + data.length + topLine);

	fillListStoresTable(data);
	
	$(".listStoresContainer").show();
	
	$(".storesColWideLnk1").click( function() { remoteConnect($("#" + this.id).text(), MFS1); });
	$(".storesColWideLnk2").click( function() { dialOut(this.id) });
	$(".storesColNarrowLnk1").click( function() { grepData($("#" + this.id).text()); });
	$("#masterRCCheckbox").change(function() {
		if ($(this).prop("checked")) { $(".rcCheckbox").prop("checked", true) }
		else { $(".rcCheckbox").prop("checked", false) };
	});
};


function fillListStoresTable(data) {
	$("#listStores").html("");
	$.each(data, function(i, n) {
		
		$("#listStores").append("\
			<tr class='storeRow' id='storeRow" + i +"'>\
				<td class='storesColNarrowLnk1 listedStore" + i +"' id='store" + i + "'>" + n.Store + "</td>\
				<td class='storesColWideLnk1 listedStore" + i +"' id='MFS1_" + i + "'>" + n.MFS1 + "</td>\
				<td class='storesColWide listedStore" + i +"'>" + n.City + "</td>\
				<td class='storesColNarrow listedStore" + i +"'>" + n.State + "</td>\
				<td class='storesColWideLnk2 listedStore" + i +"' id='phone" + i + "'>" + n.Phone + "</td>\
				<td class='storesColNarrow'><input type='checkbox' class='rcCheckbox' value='" + i +"'></td> \
			</tr>\
		");
	});
}


function listStoresMassConnect() {	// collects IP addresses of store checked in the listStores table and passes each one to the remoteConnect() function
	var storesArray = []
	$(".rcCheckbox:checked").map(function(i) {
		let rowNumber = $(this).val()
		storesArray[i] = $("#MFS1_" + rowNumber).text();
	});

	if (storesArray.length > 30) {
		if (confirm("You want to connect to " + storesArray.length + " stores?")) {
			alert("If you say so. Say good-bye to Internet Explorere for me.");
		} else {
			return;
		};
	};

	storesArray.forEach(function(i) { remoteConnect(i, "MFS1") });
}


function exportListStores() {	// exports data in the listStores to csv format for download.
	
	var csvString = "Store, Server, City, State, Phone \r\n";

	$("#listStores tr").each(function(i, r) {
		$("#storeRow" + i + " > td").each(function(ii, dd) {
			if (ii > 0 ) { csvString += ","  };
			csvString += dd.textContent;
		});
		csvString += "\r\n";
	});

	let exportData = new Blob([csvString], {type: "text/csv;charset=utf-8,"});
	navigator.msSaveBlob(exportData, "Store List Export.csv");
}


function allStateSelector(loadedState) {	// called by writeListStores function. returns the list of options for the Select field containing a list of states and territories with Company stores.

	var stateArray = [];
	var stateString = "";
	
	storeData.forEach(function(i) {	
		if (stateArray.indexOf(i["State"]) < 0 && i["State"] != "" && i["State"] != undefined){
			stateArray.push(i["State"])
		};
	});
	stateArray.sort();
	
	stateArray.forEach(function(s) {
		stateString = stateString + "<option value='" + s + "'>" + s + "</option>"
	});

	stateString = stateString.replace("value='" + loadedState +"'", "value='" + loadedState +"' selected");
	return stateString;
};


function allRegionSelector(loadedRegion) {	// called by writeListStores function. returns the list of options for the Select field containing a list of regions Company stores.
	
	var regionArray = [];
	var regionString = "";
	
	storeData.forEach(function(i) {
		if (regionArray.indexOf(i["Region"]) <0 && i["Region"] != "" && i["Region"] != undefined) {
			regionArray.push(i["Region"]);
		};
	});
	regionArray.sort();
	
	regionArray.forEach(function(r) {
		regionString = regionString + "<option value'" + r + "'>" + r + "</option>"
	});
	
	regionString = regionString.replace("value='" + loadedRegion + "'", "value='" + loadedRegion + "' selected");
	return regionString;
}


function allDCSelector(loadedDC) {	// called by writeListStores function. returns the list of options for the Select field containing a list of distribution centers for Company stores.
	
	var DCArray = [];
	var DCString = "";
	
	storeData.forEach(function(i) {
		if (DCArray.indexOf(i["DC"]) <0 && i["DC"] != "" && i["DC"] != undefined) {
			DCArray.push(i["DC"]);
		};
	});
	
	DCArray.forEach(function(d, n) {	// Adding leading 0's to DC numbers so the values will sort correctly.
		if(d.substr(1,1) == " ") {
			DCArray[n] = "0" + d
		};
	});

	DCArray.sort();

	DCArray.forEach(function(d, n) {	// Removing those leading 0's so the grep in the doListStores function will work.
		if(d.substr(0,1) == "0") {
			DCArray[n] = d.slice(1);
		};
	});
	
	DCArray.forEach(function(d) {
		DCString = DCString + "<option value'" + d + "'>" + d + "</option>"
	});
	
	DCString = DCString.replace("value='" + loadedDC+ "'", "value='" + loadedDC + "' selected");
	return DCString;
}


function sortTable(n) {
// sorts entries in the List Stores By... table.
// accepts one arguement: n - collumn number to sort by
	
	var stTable = $("#listStores");
	var sortvalue = 'td:nth-child(' + n + ')';
	var rows
	var i, ii, iii
	var y, v1, v2, cc, d

	for (i = 0; i < stTable.length; i++) {
		for (ii = 0; ii < 2; ii++) {
			cc = 0;
			y = 1;
			d = 0

			rows = stTable[i].querySelectorAll(".storeRow");
			if (rows.length > 75) {
				alert ("Sort Length Too Long.");
				break;	//Attempting to sort over 100ish rows in this manor will cause the script to run too long and freeze the web page.
			};
			while (y == 1) {
				y = 0;
				rows = stTable[i].querySelectorAll(".storeRow");
				for (iii = 0; iii < (rows.length -1); iii++) {
					v1 = rows[iii].querySelector(sortvalue).innerHTML;
					v2 = rows[iii + 1].querySelector(sortvalue).innerHTML;

					if (isNaN(v1)) {
						v1 = v1.toLowerCase();
						v2 = v2.toLowerCase();
					} else {
						v1 = parseInt(v1);
						v2 = parseInt(v2);
					}

					if (ii == 0 && (v1 > v2)) {
						rows[iii].parentNode.insertBefore(rows[iii + 1], rows[iii]);
						d = 1
						y = 1;
						cc++;
					} else if (ii == 1 && (v1 < v2)) {
						rows[iii].parentNode.insertBefore(rows[iii + 1], rows[iii]);
	
						y = 1;
						d = 2;
						cc++;
					};
				};
			};
			$(".sortTable").html("<img src='./graphic/arrow3.png'>");
			if (d == 1) { $("#sortTable" + n).html("<img src='./graphic/arrow5.png'>") }
			else if (d == 2) { $("#sortTable" + n).html("<img src='./graphic/arrow4.png'>") };
			if (cc > 0) {break;}
		}
	}
};
 
 
function movePopOut(head, elmt) {	// called on page load. Listens for the header element of the popout section to be clicked and moves the element with the mouse.

	var dX = 0;
	var oX = 0;
	var wY;
	var dY = 0;
	var oY = 0;
	var wY;

	$(head).mousedown(function () {
		moveElement();

		function moveElement(w) {
			w = w || window.event;
			oX = w.clientX;
			oY = w.clientY;

		document.onmouseup = stopMove;
		document.onmousemove = doMove;
		};

		function doMove(w) {
			w = w || window.event;
			dX = oX - w.clientX;
			dY = oY - w.clientY;
			oX = w.clientX;
			oY = w.clientY;
			wX = $(elmt).offset().left
			wY = $(elmt).offset().top
			$(elmt).offset({top: wY-dY, left: wX-dX});
		};

		function stopMove() {
			document.onmouseup = null;
			document.onmousemove = null;
		};
	});
};


function launchRemoteOptions() {
// called by event listener in document.ready. Toggles visibility of connect immediately options.
	if ($("#remote").is(":checked")) {
		$("#getInfo").html("Lookup & Connect")
		$(".immediateConnectOptions").show();
	} else {
		$("#getInfo").html("Retrieve Store Info")
		$(".immediateConnectOptions").hide();
	};
}


function clearMid() {	// hides pop-out frame in the center of the screen, or hides email templates and phone list.
	if ($("#phoneList").text().search(/show/i) == -1 || $("#emailTemplates").text().search(/show/i) == -1) {
		if($("#phoneList").text().search(/show/i) == -1) { togglePopInSections("phoneList") };
		if($("#emailTemplates").text().search(/show/i) == -1) { togglePopInSections("emailTemplates") };
	} else if ($(".popOut").is(":visible")) { 
		clearEmail();
	} else if ($(".remoteOptions").is(":visible")) {
		toggleSections("remoteOptions");
	};
}


function writeMassConnect() {	// Draws input fields for remote connect feature.
	var remoteOption = $("input[name='connectOption']:checked").val()

	if ( $(".popOut").is(":visible") && $("#popOutTitle").text() == "Mass Remote Connect" ) {
		clearEmail();
		return;
	};

	$("#popOutTitle").text("Mass Remote Connect");
	$(".popOutBody").html("\
		<div class='massConnect'>\
			Store list :<br />\
			<textarea id='massConnectList' rows='10' cols='35'></textarea>\
		</div>\
		<div class='massConnect'>\
			Connect to :\
			<p>\
			<input type='radio' name='massConnectDevice' value=MFS1 id='massConnectDevice0' checked /><label for='massConnectDevice0'> &nbsp; &nbsp; PC</label> &nbsp; &nbsp; \
			<input type='radio' name='massConnectDevice' value=MIO_PC id='massConnectDevice1' /><label for='massConnectDevice1'> &nbsp; &nbsp; MIO</label>\
			</p>\
			<input type='radio' name='massConnectOption' value=0 id='massConnectOption0' /><label for='massConnectOption0'> &nbsp; &nbsp; Dashboard</label><br />\
			<input type='radio' name='massConnectOption' value=1 id='massConnectOption1' /><label for='massConnectOption1'> &nbsp; &nbsp; Remote Control</label><br />\
			<input type='radio' name='massConnectOption' value=2 id='massConnectOption2' /><label for='massConnectOption2'> &nbsp; &nbsp; File Transfer</label><br />\
			<input type='radio' name='massConnectOption' value=3 id='massConnectOption3' /><label for='massConnectOption3'> &nbsp; &nbsp; Command Prompt</label><br />\
			<input type='radio' name='massConnectOption' value=4 id='massConnectOption4' /><label for='massConnectOption4'> &nbsp; &nbsp; Event Viewer</label><br />\
		<button onclick='doMassConnect()'>Connect</button>\
		</div>\
	");
	
	if ($(".remoteOptions").is(":visible")) { toggleSections("remoteOptions") };
	$("#massConnectOption" + remoteOption).prop("checked", true);
	$(".popOut").show();
	$("#massConnectList").focus();

	$("input[name='massConnectOption']").change(function() {
		remoteOption = $("input[name='massConnectOption']:checked").val();
		$("#connectOption" + remoteOption).prop("checked", true);	
	});
}


function doMassConnect() {	// Reads input from text field, converts to an array, confirms all items in the array are valid store numbers, then reads remote connection options and passes arguments to remoteConnect function.
	var allGood = 1;
	var connectDevice = $("input[name='massConnectDevice']:checked").val();
	var connectOption = $("input[name='massConnectOption']:checked").val();
	var listString = $("#massConnectList").text();
	var listArray = [];
	var addressArray = [];

	listString = listString.replace(" ","").replace(/\r|\n|\r\n/g,',');
	if (listString.charAt(listString.length - 1) == " ") { listString = listString.slice(0,-1) };;

	listArray = listString.split(",");

	for (var n = 0; n < listArray.length; n++) {
		data = $.grep(storeData, function(line) {
			return (line.Store == listArray[n] || line.Orical == listArray[n]);
		});

		try {
			if (isNaN(listArray[n]) || listArray[n] == "") throw listArray[n] + " is not a vaild store number.";
			if (data[0] === undefined) throw listArray[n] + " is not a vaild store number.";
			if (data[0]["CPcompany"] != 115741 && connectDevice == "MIO_PC") throw listArray[n] + " does not have a MIO PC."
		} catch(erMsg) {
			alert( erMsg + " \nEnter one store number per line." );
			allGood = 0
			break;
		};
		addressArray.push(data[0][connectDevice]);
	};
	if (allGood) { addressArray.forEach(function(i) { remoteConnect(i, connectDevice) }) };
};