//	Each email template has two functions controlling them. the write[X] function template draws the input fields and shows the popOut element. The send[X] function checks to make sure required information is present before passing the email parameters to the sendMail function to click the link and generate the email. 

function sendMail(To, Subject, Body, Copy) {
	if (confirm("This information is correct?")) {
		if (Copy === undefined) {
			var Email = "mailto:" + To + "&subject=" + Subject + "&body=" + Body;
		} else {
			var Email = "mailto:" + To + "?cc=" + Copy + "&subject=" + Subject + "&body=" + Body;
		};
		clearEmail();
		window.location.href = Email;
	} else {
		return;
	};	
};

function clearEmail() {
		$(".popOutBody").text("");
		$("#popOutTitle").text("");
		$(".popOut").hide();	
};

function loadPartsList(data) {
	Papa.parse(data, {
		delimiter: ',',
		quoteChar: '"',
		header: true,
		complete: function(results) {
			console.log("Parts list loaded and parsed:", results);
			partsList = results.data;
		}
	});
};

function writeEmail(template) {
	//	Determines which template to display.
	switch (template) {
		case 1:
			writeCPInfoUpdate();
			break;
		case 2:
			writeDUD();
			break;
		case 3:
			writeEmailStore();
			break;
		case 4:
			writeCPPasswordReset();
			break;
		case 5:
			writeOrderParts();
			break;
		case 6:
			writeServiceNetworkDrop();
			break;
		case 7:
			writeRequestTimeclock();
			break;
		case 8:
			writeScaleCalibration();
			break;
		case 9:
			writeResetCounters();
			break;
		case 10:
			writeSBOInstalled();
			break;
		case 11:
			writeSBOSetup();
			break;
		case 12:
			writeScanbackEligibility();
			break;
		case 13:
			writeServerSwap();
			break;
		case 14:
			writeLabCallout()
			break;
		case 15:
			writeCallTags()
			break;
	}
	togglePopInSections("emailTemplates");
}

function sendEmail() {	//	Called by eventFire() function. Determines which send[X] function to call based on the text in the popOutTitle element
	switch ($("#popOutTitle").text()) {
		case "Request Scanner/scale Calibration":
			sendScaleCalibration();
			break;
		case "Request MicroTrax Password Reset":
			sendCPPasswordReset();
			break;
		case "CP Info Update":
			sendCPInfoUpdate();
			break;
		case "Request New Timeclock":
			sendRequestTimeclock();
			break;
		case "Reset Counters":
			sendResetCounters();
			break;
		case "Order POS Parts":
			sendOrderParts();
			break;
		case "Email Store":
			sendEmailStore();
			break;
		case "Request RDS to Service Network Drop":
			sendServiceNetworkDrop();
			break;
		case "SBO Setup - Change Setup Scripts":
			sendSBOSetup();
			break;
		case "SBO Setup - Installation Complete":
			sendSBOInstalled();
			break;
		case "DUD Equipment to Store":
			sendDUD();
			break;
		case "Server Swap Notification":
			sendServerSwap();
			break;
		case "Call Tags Email":
			sendCallTags();
			break;
	};
};

function formatDate(t) {
	var today = new Date();
	var dt = today.getDate();
	var mo = today.getMonth()+1;
	var yr = today.getFullYear();
	var dy = today.getDay();
	var date
	
	if (dt < 10) { dt = "0" + dt };
	if (mo < 10) { mo = "0" + mo };
	switch (t) {
		case 1:
			date = mo + "/" + dt + "/" + yr;
			break;
		case 2:
			date = mo + "/" + dt;
			break;
	};
		
	return date;
};

function sanitize(input) {
	input = input.replace(/\%/g, '%25').replace(/\r|\n|\r\n/g,'%0d%0a').replace(/#/g, '%23').replace(/&/g, '%26').replace(/"/g, '%22');
	return input;
}

function laneCount(storeNum) {	//	looks up the number of lanes of the store number passed to the function. defaults to 4 lanes if information is unavailable.

	var company
	var counter = 1;
	var list = ["<option selected disabled value='null'></option>"];
	var times
	
	if ($.isNumeric(storeNum)) {	
		var data = $.grep(storeData, function(line) {
			return (line.Store == storeNum || line.Orical == storeNum);
		});
		if(data[0] === undefined) { 
			alert( storeNum + " is not a valid store number." );
			$("#emailField1").val("").focus();
			times = 4;
		} else {
			times = data[0].NumLanes;
			company = data[0].CPcompany;
		};
	} else {
		times = 4;	
	};
	
	for (; counter <= times; counter++) { list += "<option value='Lane " + counter + "'>Lane " + counter + "</option>" };

	return [list, company];
};

function writeScaleCalibration() {
	var loadedStore = $("#storeNumber").text();
	var lanes = laneCount(loadedStore)

	$("#popOutTitle").text("Request Scanner/scale Calibration");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5' > &nbsp; &nbsp; Lane : <select id='emailField2'>" + lanes[0] + "</select> &nbsp; &nbsp; <span>S/N : <input id='emailField4' type='text' maxlength=11 size='11'></span></p>\
	<p>Reason : \
	<input type='radio' name='reason' id='SSreplace' value='1'><label for='SSreplace'>Replacement ordered</label> &nbsp; &nbsp; \
	<input type='radio' name='reason' id='SStagged' value='2'><label for='SStagged'>Tagged by Weights & Measures</label> &nbsp; &nbsp; \
	<input type='radio' name='reason' id='SSreplace' value='3'><label for='SSreplace'>Other (specify)</label> &nbsp; &nbsp; \
	<br><span id='otherReason'>Reason : <input id='emailField5' type='text' size='50'></span>\
	<span id='RMA'>RMA # <input id='emailField3' type='text' maxlength=11 size='11'></span>\
	</p><button onclick='sendScaleCalibration()'>C<u>r</u>eate Email</button>");
	$("#RMA").hide();
	$("#otherReason").hide();
	$(".popOut").show();
	$("#emailField1").focus();
	
	$("input[name='reason']").change(function() {	// toggle the RMA field to show only when the correct radio button is checked
		switch ($("input[name='reason']:checked").val()) {
			case "1":
				$("#RMA").show();
				$("#otherReason").hide();
				break;
			case "3":
				$("#otherReason").show();
				$("#RMA").hide();
				break;
			default:

				$("#RMA").hide();
				$("#otherReason").hide();
				break;
		};
	});
	
	$("#emailField1").change(function() {	// update lane count when store number field changes
		lanes = laneCount($("#emailField1").val());
		$("#emailField2").empty().append(lanes[0])
	});
};

function sendScaleCalibration() {
	var Store = $("#emailField1").val();
	var Lane = $("#emailField2").val()
	var Radios = $("input[name='reason']:checked").val();
	var RMA = $("#emailField3").val();
	var serialNumber = $("#emailField4").val();
	var otherReason = $("#emailField5").val();
	var Reason = ""
	try {
		if (Store == "") throw "Please enter a store number";
		if (isNaN(Store)) throw "Please enter a numeric store number";
		if (Lane == null) throw "Please select a lane";
		if (serialNumber == "") throw "Serial numbers are now required to request scanner/scale calibration";
		if (Radios === undefined) throw "Please select a reason";
		if (Radios == 1 && RMA == "") throw "Please enter the RMA #";
		if (Radios == 3 && otherReason == "") throw "Please specify an other reason";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	Lane = Lane.toLowerCase();
	
	switch (Radios) {
		case "1":
			Reason = "is being replaced in RMA " + encodeURIComponent("#") + " " + RMA;
			break;
		case "2":
			Reason = "was tagged by weights and measures";
			break;
		case "3":
			Reason = "needs calibration for the following reason: " + otherReason;
			break;
	};

	var To = "Brenda@Company.com";
	var Copy = "Jay@Company.com; BJ@Company.com; Dusty@Company.com;";
	var Subject = "Scale Calibration : Company " + Store + " " + Lane;
	var Body = "Hi Brenda%0d%0a%0d%0aThe scanner/scale on " + Lane + " (serial number " + serialNumber + ") at store " + Store + " " + Reason + ". Could you schedule a tech to calibrate it" + encodeURIComponent("?") + "%0d%0a%0d%0aThank you.";
	sendMail(To, Subject, Body, Copy);
};

function writeCPPasswordReset() {
	var loadedStore = $("#storeNumber").text();
	var loadedCPUser = $("#cpuser").text();
	var loadedCPNumber = $("#cpcompany").text();
	$("#popOutTitle").text("Request MicroTrax Password Reset");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'> &nbsp; &nbsp; Company ID : <input id='emailField2' type='text' value='" + loadedCPNumber + "' maxlength=6 size='6'>  &nbsp; &nbsp; User Name : <input id='emailField3' tye='text' value='" + loadedCPUser + "'></p><button onclick='sendCPPasswordReset()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
};

function sendCPPasswordReset() {
	var Store = $("#emailField1").val();
	var Company = $("#emailField2").val();
	var User = sanitize($("#emailField3").val());

	try {
		if (Store == "") throw "Please enter a store number";
		if (isNaN(Store)) throw "Please enter a numeric store number";
		if (Company == "") throw "Please enter a company number";
		if (isNaN(Company)) throw "Please enter a numeric company number";
		if (User === undefined) throw "Please enter a user name";
		if (User == "") throw "Please enter a user name";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	var To = "Support@Vendor.com";
	var Subject = "ServerEPS Password reset : Company " + Store;
	var Body = "Please reset the password for Company store " + Store + "%0d%0aUsername : " + User + "%0d%0aCompany no : " + Company;
	sendMail(To, Subject, Body);
};

function writeCPInfoUpdate() {
	var loadedCPNumber = $("#cpcompany").text();
	$("#popOutTitle").text("CP Info Update");
	$(".popOutBody").html("<p>Company ID : &nbsp;&nbsp;<input id='emailField1' type='text' value='" + loadedCPNumber + "' maxlength=6 size='6'></p><p>New password : <input id='emailField2' type='text'>  &nbsp; &nbsp; Confirm password : <input id='emailField3' type='text'></p><button onclick='sendCPInfoUpdate()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
};

function sendCPInfoUpdate() {
	var Company = sanitize($("#emailField1").val());
	var iPassword = sanitize($("#emailField2").val());
	var cPassword = sanitize($("#emailField3").val());
	
	try {
		if (Company == "") throw "Please enter a company number";
		if (isNaN(Company)) throw "Please enter a numeric company number";
		if (iPassword === undefined) throw "Please enter the new password";
		if (cPassword === undefined) throw "Please confirm the new password";
		if (iPassword === "") throw "Please enter the new password";
		if (cPassword === "") throw "Please confirm the new password";
		if (iPassword != cPassword) throw "Passwords do not match";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	var To = "Larry@Company.com";
	var Copy = "SupportTeam@Company.com";
	var Subject = "CP Info Update - co" + encodeURIComponent("#") + Company;
	var Body = "Hello Rick,%0d%0a%0d%0aPlease change the connected payments password for company " + Company + " to " + iPassword + "%0d%0a%0d%0aThanks.";
	sendMail(To, Subject, Body, Copy);
};

function writeRequestTimeclock() {
	var loadedStore = $("#storeNumber").text();
	var loadedAddress1 = $("#storeAddress").text();
	var loadedAddress2 = $("#storeCity").text() + ", " + $("#storeState").text() + " " + $("#storeZip").text()
	$("#popOutTitle").text("Request New Timeclock");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'></p><p>Address :<br><textarea id='emailField2' rows='2' cols='45'>" + loadedAddress1 + "&#13&#10" + loadedAddress2 + "</textarea></p><p><input type='radio' name='model' id='radio1' value='1'><label for='radio1'>Kronos 4500</label> &nbsp; &nbsp; <input type='radio' name='model' id='radio2' value='2'><label for='radio2'>Kronos inTouch</label></p><p>Biometric S/N : &nbsp;&nbsp;<input id='emailField3' type='text'><br>Time Clock S/N: <input id='emailField4' type='text'></p><p>Reason :<br><textarea id='emailField5' rows='6' cols='45'></textarea></p><button onclick='sendRequestTimeclock()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
};

function sendRequestTimeclock() {
	var Store = $("#emailField1").val();
	var Address = sanitize($("#emailField2").val());
	var bmSN = $("#emailField3").val();
	var tcSN = $("#emailField4").val();
	var Reason = sanitize($("#emailField5").val());
	var Radios = $("input[name='model']:checked").val();
	
	try {
		if (Store == "") throw "Please enter a store number";
		if (isNaN(Store)) throw "Please enter a numeric store number";
		if (Store > 9001) throw "Time clocks are only ordered for corporate stores";
		if (Address.length < 3) throw "Please enter the store address";
		if (Radios === undefined) throw "Please select a model";
		if (bmSN == "" && Radios != 2) throw "Please enter the serial number of the biometric device";
		if (tcSN == "") throw "Please enter the serial number of the time clock";
		if (Reason == "") throw "You must give a reason for requesting a replacement time clock";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	if (Radios == 1) {
		var Model = "Kronos 4500";
		var Serial = "Biometric S/N - " + bmSN + "%0d%0a" + Model + " S/N - " + tcSN;
	} else if (Radios == 2) {
		var Model = "Kronos inTouch";
		if (bmSN == "" ) {
			var Serial = Model + " S/N - " + tcSN;
		} else {
			var Serial = "Biometric S/N - " + bmSN + "%0d%0a" + Model + " S/N - " + tcSN;
		};
	};

	var To = "Ron@Vendor.com";
	var Subject = "Company " + Store + " needs a new " + Model + " time clock";
	var Body = "Hi Ron,%0d%0a%0d%0aWe're having problems with one of our time clocks%0d%0a" +Reason + "%0d%0aPlease send a new " + Model + "%0d%0a%0d%0a" + Serial + "%0d%0a%0d%0aShip to :%0d%0aAttn: store manager%0d%0a" + Address;
	sendMail(To, Subject, Body);
};

function writeResetCounters() {
	var loadedStore = $("#storeNumber").text();
	var loadedCPNumber = $("#cpcompany").text();
	var loadedOldCPNumber = $("#storeCPOld").text();
	$("#popOutTitle").text("Reset Counters");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'> &nbsp; &nbsp; Company ID : <input id='emailField2' type='text' value='" + loadedCPNumber + "' maxlength=6 size='6'> &nbsp; &nbsp; Env. 7 Company ID : <input id='emailField3' type='text' value='" + loadedOldCPNumber + "' maxlength=6 size='6'></p><button onclick='sendResetCounters()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
};

function sendResetCounters() {
	var Store = $("#emailField1").val();
	var env2Number = $("#emailField2").val();
	var env7Number = $("#emailField3").val();
	
	try {
		if (Store == "") throw "Please enter the store number";
		if (isNaN(Store)) throw "Please enter the numeric store number";
		if (env2Number == "") throw "Please enter the company number";
		if (isNaN(env2Number)) throw "Please enter the numeric company number";
		if (env7Number == "") throw "Please enter the old company number";
		if (isNaN(env7Number)) throw "Please enter the old numeric company number";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	var To ="Support@Vendor.com";
	var Subject = "MFS server down server swap for store " + Store + " of company " + env2Number + " (" + env7Number + ")";
	var Body = "Please reset the counters for store " + Store + " of company " + env2Number + " (" + env7Number + ") and confirm good data communication.";
	sendMail(To, Subject, Body);
};

function writeOrderParts() {
	
	var loadedStore = $("#storeNumber").text();
	$("#popOutTitle").text("Order POS Parts");
	$(".popOutBody").html("<p>Tech Name: <input id='emailField1' type='text' value='" + techName.replace(/'/,'&#39;') + "' size='16'>&nbsp; &nbsp;Store: <input id='emailField2' type='text' value='" + loadedStore + "' maxlength=5 size='5'> &nbsp; &nbsp; Shipping: <input type='radio' name='shipping' id='ShipG' value='Ground' checked><label for='ShipG'>Ground</label> &nbsp; <input type='radio' name='shipping' id='ShipO' value='Overnight'><label for='ShipO'>Overnight</label> &nbsp; <input type='radio' name='shipping' id='ShipA' value='Overnight AM'><label for='ShipA'>AM</label></p><div class='partText'>Equipment Description : <br><textarea id='emailField3' rows='3' cols='37'></textarea></div><div class='partText'>Special Instructions :<span class='isHouchens'><input type='checkbox' id='isHouchens'><label for='isHouchens'>Houchens store</label></span><br><textarea id='emailField4' rows='3' cols='37'></textarea></div><div class='popOutParts'></div>");
	$(".popOutParts").append("<div class='partHeader' id='HP'><img src='./graphic/arrow.png' id='HPArrow' class='arrow'> HP Parts</div><div class='partsContainer' id='HPPartsContainer'><div class='partsList' id='HPParts'></div></div>");
	$(".popOutParts").append("<div class='partHeader' id='NCR'><img src='./graphic/arrow.png' id='NCRArrow' class='arrow'> NCR Parts</div><div class='partsContainer' id='NCRPartsContainer'><div class='partsList' id='NCRParts'></div></div>");
	$(".popOutParts").append("<div class='partHeader' id='Misc'><img src='./graphic/arrow.png' id='MiscArrow' class='arrow'> Miscellaneous Parts</div><div class='partsContainer' id='MiscPartsContainer'><div class='partsList' id='MiscParts'></div></div>");
	$(".popOutBody").append("<button onclick='sendOrderParts()'>C<u>r</u>eate Email</button>");
	var h = 0;
	var he = 0;
	var c = 0;
	var ce = 0;
	var m = 0;
	var me = 0;
	var type
	var subtype
	$.each(partsList, function(i, n) {	// goes down the list of entries in the partsList and places each with a checkbox in the correct DOM based on the Model and Subtype.
		if (n.Model == "HP") {
			if (type != n.Type ) {
				type = n.Type;
				h = 0;				
				$("#HPParts").append("<div class='partType'>" + type + "</div>");	
			};
			
			if (n.Subtype != "" && subtype != n.Subtype) {
				subtype = n.Subtype;
				h = 0;				
				$("#HPParts").append("<div class='partSubType'>" + subtype + "</div>");	
			};

			$("#HPParts").append("<div class='part'><input type='checkbox' class='partsCheckbox' value=' " + n.Number + "' id='partNo" + i + "' /><label for='partNo" + i + "'>" + n.Description + "</label></div>");
				h = h + 1;
				he = h % 2;
				if (he == 0) {
					$("#HPParts").append("<br />");
				};
			
		} else if (n.Model == "NCR") {
			if (type != n.Type ) {
				type = n.Type;
				c = 0;				
				$("#NCRParts").append("<div class='partType'>" + type + "</div>");	
			};

			if (n.Subtype != "" && subtype != n.Subtype) {
				subtype = n.Subtype;
				c = 0;				
				$("#NCRParts").append("<div class='partSubType'>" + subtype + "</div>");	
			};
			
			$("#NCRParts").append("<div class='part'><input type='checkbox' class='partsCheckbox' value=' " + n.Number + "' id='partNo" + i + "' /><label for='partNo" + i + "'>" + n.Description + "</label></div>");
			c = c +1
			ce = c % 2
			if (ce == 0) {
				$("#NCRParts").append("<br />");
			};
		} else if (n.Model == "Misc") {

			if (type != n.Type ) {
				type = n.Type;
				m = 0;				
				$("#MiscParts").append("<div class='partType'>" + type + "</div>");	
			};
			
			if (n.Subtype != "" && subtype != n.Subtype) {
				subtype = n.Subtype;
				m = 0;				
				$("#MiscParts").append("<div class='partSubType'>" + subtype + "</div>");	
			};
			
			$("#MiscParts").append("<div class='part'><input type='checkbox' class='partsCheckbox' value=' " + n.Number + "' id='partNo" + i + "' /><label for='partNo" + i + "'>" + n.Description + "</label></div>");
			m = m + 1
			me = m % 2
			if (me == 0) {
				$("#MiscParts").append("<br />");
			};
		};
	});

	if ($("#retailer").text() == " - Houchens") { $("#isHouchens").prop("checked", true); };
	
	$(".partsContainer").hide()
	$(".popOut").show();
	$("#emailField1").focus();
	
	$(".partHeader").click(function(id) {
		var container = "#" + this.id + "PartsContainer"
		var arrow = "#" + this.id + "Arrow"
		if ($(container).is(":hidden")) {
			$(".partsContainer").hide()
			$(".arrow").css("transform","");
		};
		$(container).toggle();
		
		if ($(arrow).css("transform") === "matrix(0, 1, -1, 0, 0, 0)" ){ 
			$(arrow).css("transform","");
		}else{
			$(arrow).css("transform", "rotate(90deg)");
		};
	
	});
};

function sendOrderParts() {
	var Tech = sanitize($("#emailField1").val());
	var Store = $("#emailField2").val();
	var Description = sanitize($("#emailField3").val());
	var Instructions = sanitize($("#emailField4").val());
	var Today = formatTime(3);
	var Items = []
	var Shipping = $("input[name='shipping']:checked").val();


	$(".partsCheckbox:checked").map(function(i) {
		Items[i] = $(this).val();
	});

	var List = Items.toString();
	
	try {
		if (Tech == "") throw "Please enter your name";
		if (Store == "") throw "Please enter a store number";
		if (isNaN(Store)) throw "Please enter a numeric store number";
		if (Description == "") throw "Please enter the item(s) description for confirmation.";
		if (List.indexOf(89184) >0 ^ List.indexOf(89187) >0) throw "The power cable & power brick for the 9590 must be ordered together.";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};

	if (List == []) {
		List = "Unknown item number"
	};

	if ($("#isHouchens").is(":checked")) { Instructions = Instructions + " This is for a Houchens store." };

	var To = "Brenda@Company.com"
	var Copy = "Dusty@Company.com; Tom@Company.com; SupportTeam@Company.com"
	var Subject = "Tier 2 Support Equipment Request"
	var Body = "Tech name : " + Tech + "%0d%0a%0d%0aToday's date : " + Today + "%0d%0a%0d%0aStore Number : " + Store + "%0d%0a%0d%0aShipping Method : " + Shipping + "%0d%0a%0d%0aItem Number(s) : " + List + "%0d%0a%0d%0aDescription : " + Description + "%0d%0a%0d%0aSpecial Instructions : " + Instructions
	sendMail(To, Subject, Body, Copy);
	techName = Tech;
	document.cookie = "techName = " + Tech + ";";
	
	
};

function writeEmailStore() {
	var loadedStore = $("#storeNumber").text();
	$("#popOutTitle").text("Email Store");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'> </p><button onclick='sendEmailStore()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
};

function sendEmailStore() {
	var store = $("#emailField1").val();
	var address
	
	try {
		if (store == "") throw "Please enter the store number";
		if (isNaN(store)) throw "Please enter a numeric store number";
		if (store.length > 5 || store.length < 3) throw "Valid store numbers are between 3 and 5 digits";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	switch (store.length) {
		case 3:
			address = "Retail.Store.00" + store + "@store.company.com";
			break;
		case 4:
			address = "Retail.Store.0" + store + "@store.company.com";
			break;
		case 5:
			address = "Retail.Store." + store + "@store.company.com"
			break;
	};

	var To = address ;
	var Subject = " ";
	var Body = " ";
	sendMail(To, Subject, Body);
};

function writeServiceNetworkDrop() {
	var loadedStore = $("#storeNumber").text();
	var dropList = makeDropList(loadedStore);
	
	$("#popOutTitle").text("Request RDS to Service Network Drop");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'> &nbsp; &nbsp; Affected equipment : <select id='emailField2'>" + dropList + "</select></p><p>Reason :<br /><textarea id='emailField3' rows='6' cols='45'></textarea></p><button onclick='sendServiceNetworkDrop()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
	
	$("#emailField1").change(function() {	// update lane count when store number field changes
		dropList = makeDropList($("#emailField1").val());
		$("#emailField2").empty().append(dropList)
	});
	
	function makeDropList(storeNum) {
		var info = laneCount(storeNum)

		if (info[1] != 115741) {
			drops = info[0] + "<option value='the access point'>Access Point</option><option value='the Financial Center'>Financial Center</option><option value='a price checker'>Price checker</option>";
		} else {
			drops = info[0] + "<option value='the access point'>Access Point</option><option value='the meat scale'>Meat Scale</option><option value='the Mio PC'>Mio PC</option><option value='the produce scale'>Produce Scale</option><option value='a price checker'>Price checker</option><option value='the safe'>Safe</option><option value='the time clock'>Time Clock</option>";
		};
		return drops;
	};
};
	
function sendServiceNetworkDrop() {
	var Store = $("#emailField1").val();
	var Equipment = $("#emailField2").val();
	var Reason = sanitize($("#emailField3").val());
	
	try {
		if (Store == "") throw "Please enter the store number";
		if (isNaN(Store)) throw "Please enter the numeric store number";
		if (Equipment == null) throw "Please select the equipment to be serviced";
		if (Reason == "") throw "Please describe the work to be done.";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	var To = "BJ@Company.com";
	var Copy ="Jay@Company.com";
	var Subject = "Request for Network Cable Maintenance : Company " + Store;
	var Body = "BJ,%0d%0a%0d%0aCould you arrange for a tech from Direct Source to service the network drop for " + Equipment + " at store " + Store + "?%0d%0a%0d%0a" + Reason + "%0d%0a%0d%0aThank you.";
	sendMail(To, Subject, Body, Copy);
};

function writeSBOSetup() {
	var loadedStore = $("#storeNumber").text();
	$("#popOutTitle").text("SBO Setup - Change Setup Scripts");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'> </p><button onclick='sendSBOSetup()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
};

function sendSBOSetup() {
	var Store = $("#emailField1").val();
	
	try {
		if (Store == "") throw "Please enter the store number";
		if (isNaN(Store)) throw "Please enter the numeric store number";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	var To ="BJ@Company.com; Trent@Company.com";
	var Subject = "SBO Install - " + Store;
	var Body = "Please change the setup script for store " + Store + ".%0d%0a%0d%0aThanks.";
	sendMail(To, Subject, Body);
};

function writeSBOInstalled() {
	var loadedStore = $("#storeNumber").text();
	$("#popOutTitle").text("SBO Setup - Installation Complete");
	$(".popOutBody").html("<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'> </p><button onclick='sendSBOInstalled()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
};

function sendSBOInstalled() {
	var Store = $("#emailField1").val();
	
	try {
		if (Store == "") throw "Please enter the store number";
		if (isNaN(Store)) throw "Please enter the numeric store number";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	var To ="BJ@Company.com; Trent@Company.com; Jay@Company.com; Tim@Company.com";
	var Subject = "SBO Install - Store " + Store + " complete";
	var Body = "Installation of the replacement SBO computer at store " + Store + " is complete. Please update the shared document.%0d%0a%0d%0aThanks.";
	sendMail(To, Subject, Body);
};

function writeDUD() {
	var loadedStore = $("#storeNumber").text();
	var equipmentList = equipmentList();
	var laneList = laneCount(loadedStore);
	var numRows = 0;

	$("#popOutTitle").text("DUD Equipment to Store");
	$(".popOutBody").html("<p><div class='DUDBody' id='DUDLeft'>Store &nbsp; : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='7' /><br />Ticket : <input id='emailField2' type='text' maxlength=7 size='7' /></div><div class='DUDBody' id='DUDRight'><span class='DUD'> Equipment : <select id='equip0' onchange='changeEquip(0)'>" + equipmentList + "</select> &nbsp; &nbsp; <span id='laneField0'>Lane : <select class='laneSelect' id='lane0'>" + laneList + "</option></select></span></span></div><div class='addSubButtons'><img src='./graphic/minus.png' id='subButton' /><img src='./graphic/plus.png' id='addButton' /></div></p><button onclick='sendDUD()'>C<u>r</u>eate Email</button>");
	
	$(".popOut").show();
	$("#emailField1").focus();

	$("#addButton").click(function() { addDUDRow() });
	$("#subButton").click(function() { subDUDRow() });
	
	$("#emailField1").change(function() {	// update lane count when store number field changes
		laneList = laneCount($("#emailField1").val());
		$(".laneSelect").empty().append(laneList[0]);
	});

	function equipmentList() {
		var equipmentList = "<option selected disabled value=''></option><option value='Cash Drawer'>Cash Drawer</option><option value='MIO PC'>MIO PC</option><option value='Monitor'>Monitor</option><option value='Pinpad Body'>Pinpad Body</option><option value='Pinpad I/O module'>Pinpad I/O module</option><option value='POS Printer'>POS Printer</option><option value='POST'>POST</option><option value='Price Checker'>Price Checker</option><option value='Price Checker Printer'>Price Checker Printer</option><option value='SBO PC'>SBO PC</option><option value='Server'>Server</option><option value='Touch Screen'>Touch Screen</option>";
		return equipmentList;
	};
 
	function addDUDRow() {
		numRows = numRows + 1;
		$("#DUDRight").append("<div class='DUDRow DUD' id='DUDRow" + numRows +"'>Equipment : <select class='equipField' id='equip" + numRows + "' onchange='changeEquip(" + numRows + ")'>" + equipmentList + "</select> &nbsp; &nbsp; <span id='laneField" + numRows + "'>Lane : <select class='laneSelect' id='lane" + numRows + "'>" + laneList + "</select></span></div>");
	};

	function subDUDRow() {
		if (numRows == 0) { return; };
		$("#DUDRow" + numRows).detach();
		numRows = numRows - 1;
	};

};

function changeEquip(row) {
		// function for the writeDUD function called by equipment select field to determine if the lane field needs to be shown or hidden. when embedded in the writeDUD function, the page returned an error stating that the function was undefined. I'd rather find a way for it to be in the writeDUD function, but it's working out here.
		
	var noLane = ["Server", "MIO PC", "Price Checker", "Price Checker Printer", "SBO PC"];	// which equipment doesn't require an entry in the lane field.

	if ($.inArray($("#equip" + row).val(),noLane) > -1 ) {
		$("#laneField" + row).hide();
		$("#lane" + row).val('null');
	} else {
		$("#laneField" + row).show();
	};
};

function sendDUD() {
	var d = new Date();
	var store = $("#emailField1").val();
	var ticket = $("#emailField2").val();
	var noLane = ["Server", "MIO PC", "Price Checker", "Price Checker Printer", "SBO PC"];
	var equipList = []
	var required = []
	var equipmentArray = []
	var address
	var equipment
	var weekendText =""
	
 	$(".DUD").map(function(i) {	// The first two arrays mapped are used in the try/catch error checking. The third is for creating the equipment string for the body of the email.
		equipList[i] = $("#equip" + i).val();
		if ($.inArray($("#equip" + i).val(),noLane) == -1 && $("#lane" + i).val() == null) {
			required[i] = 0;	//fail
		} else {
			required[i] = 1;	//pass
		};
		if ($.inArray($("#equip" + i).val(),noLane) > -1 ) {
			equipmentArray[i] = " a " + $("#equip" + i).val();
		} else {
			equipmentArray[i] = " a " + $("#equip" + i).val() + " for " + $("#lane" + i).val();
		};
	});
	
	try {
		if (store == "") throw "Please enter the store number";
		if (isNaN(store)) throw "Please enter a numeric store number";
		if (store.length < 3) throw "Valid store numbers are between 3 and 5 digits";
		if (ticket.search(/^INC\d{7}$/)) throw "Please enter int incident number";
		if (!equipList.every(function(n) { return n != null; })) throw "Please enter the equipment DUDed";
		if (!required.every(function(n) { return n !=0; })) throw "A lane number is required";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	equipment = equipmentArray.toString().replace(/,([^,]*)$/, " and$1");
	equipment = equipment.replace("a", "");
	equipment = equipment.replace(" ", "");

	
	switch (store.length) {
		case 3:
			address = "Retail.Store.00" + store + "@store.company.com";
			break;
		case 4:
			address = "Retail.Store.0" + store + "@store.company.com";
			break;
		case 5:
			address = "Retail.Store." + store + "@store.company.com"
			break;
	};
// Add check for date to determine if the email is being sent on a weekend. Add ' via UPS overnight shipping 
	if (d.getDay() == 0 || d.getDay() == 6) { weekendText = 'The order will be processed on Monday for delivery on Tuesday. ' };

	var To = address
	var Subject = "Attention: Store Manager";
	var Body = "Attention: Store Manager%0d%0a%0d%0aOur POS vendor, Retail Data Systems, is shipping a replacement" + equipment + " to your store via UPS overnight shipping. " + weekendText + "Please call the helpdesk and reference ticket number " + ticket + " for help on installation when the new equipment arrives, and also to make sure the new equipment is working properly.%0d%0a%0d%0aTo avoid freight charges, it is imperative you install the new equipment and return the old equipment. Even if you find later on that your old equipment is working and you do not need the new equipment, you still need to install the new equipment since it has gone through inspection and preventive maintenance and will give you better performance.%0d%0a%0d%0aIf you return the new equipment to the vendor, they will charge your store for overnight freight charges.%0d%0a%0d%0a***Note*** If the old equipment is not returned within 1 week your store will be charged for the missing equipment!%0d%0a%0d%0aThank You for Your Cooperation,%0d%0a%0d%0aRetail Systems Support%0d%0aCompany";
	sendMail(To, Subject, Body);
};

function writeServerSwap() {
	var loadedStore = $("#storeNumber").text();
	
	$("#popOutTitle").text("Server Swap Notification");
	$(".popOutBody").html("\
		<p>\
			Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'>\
			&nbsp; &nbsp; <input type='radio' name='reason' id='reason0' value='begun'><label for='reason0'>Beginning</label>\
			&nbsp; &nbsp; <input type='radio' name='reason' id='reason1' value='finished'><label for='reason1'>Finished</label>\
		</p>\
		<button onclick='sendServerSwap()'>C<u>r</u>eate Email</button>");
	$(".popOut").show();
	$("#emailField1").focus();
}

function sendServerSwap() {
	var Store = $("#emailField1").val();
	var Reason = $("input[name='reason']:checked").val();

	try {
		if (Store == "") throw "Please enter the store number";
		if (isNaN(Store)) throw "Please enter the numeric store number";
		if (Reason == undefined) throw "Please select a reason";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	var To = "SupportTeam@Company.com";
	var Subject = "Company " + Store + " server swap " + Reason;
	var Body = "Server swap for store " + Store + " has " + Reason + ".";
	sendMail (To, Subject, Body);
}

function writeScanbackEligibility() {
	var loadedStore = $("#storeNumber").text();
	var promoList = promoSelecter();
	
	$("#popOutTitle").text("Scanback Eligibility Inquirty");
	$(".popOutBody").html("\
		<p>Store : <input id='emailField1' type='text' value='" + loadedStore + "' maxlength=5 size='5'>\
		&nbsp; &nbsp; UPC : <input id='emailField2' type='text' maxlength=12 size='12'>\
		</p> <p>Description : <input id='emailField3' type='text' size='26'>\
		</p> <p>Brand : <select id='emailField4' name='promoSelect'>" + promoList + "</select>\
		&nbsp; &nbsp; <span id='otherPromo'> Please specify : <input id='emailField5' type='text'></span></p>\
		<button onclick='sendScanbackEligibility()'>C<u>r</u>eate Email</button>");
	$("#otherPromo").hide();
	$(".popOut").show();
	$("#emailField1").focus();
	
	function promoSelecter() {
		
		var brandArray = ["Coke", "Pepsi", "Frito Lay's", "Little Debbie", "Rockstar"];	// If a product becomes frequently used for Scanback promotions, add it to the selector here.
		var brandString = "";
		
		brandArray.forEach(function(b) {
			brandString = brandString + "<option value='" + b + "'>" + b + "</option>"
		});
		brandString = brandString + "<option value='Other'>Other</option>"
		
		return brandString;
	}
	
	$("#emailField4").change(function() {
		if ($("#emailField4").val() == "Other") {
			$("#otherPromo").show();
		} else {
			$("#otherPromo").hide();
		};
	});
}

function sendScanbackEligibility() {
	var Store =  $("#emailField1").val();
	var UPC =  $("#emailField2").val();
	var Description = $("#emailField3").val();
	var Promo =  $("#emailField4").val();
	var cityState;
	
	if (Promo == "Other") { Promo = $("#emailField5").val() };
	
	try {
		if (Store == "") throw "Please enter the store number";
		if (isNaN(Store)) throw "Please enter the numeric store number";
		if (UPC == "") throw "Please enter the store UPC";
		if (isNaN(UPC)) throw "Please enter the numeric UPC";
		if (Description == "") throw "Please enter the product description";
		if (Promo == "") throw "Come on, now. If you're going to select Other, you at LEAST have to tell me what it is. I'm a computer. I need to know things.";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	storeInfo = $.grep(storeData, function(line) {
		return (line.Store == Store || line.Orical == Store);
	});
	
	cityState = storeInfo[0].City + ", " + storeInfo[0].State
	Store = storeInfo[0].Store

	var To = "Alicia@Company.com";
	var Subject = Promo + " scanback item eligibility.";
	var Body = "Hi Alicia,%0d%0a%0d%0aStore " + Store + " in " + cityState + " called to see if UPC " + UPC + " (" + Description + ") is eligible for this weeks scanback promotion. Can you confirm this?%0d%0a%0d%0aThank you.";
	var CC = "Ally@Company.com";
	if (Promo == "Pepsi") { CC += "; Michael@Vendor.com" };
	sendMail (To, Subject, Body, CC);
}

function sendInLab(lab) {

	var PC = $("#" + lab).text();
	var address = $("#" + lab).attr("href")
	
	let lastDot = address.lastIndexOf(".");
	let secondColon = address.indexOf(":",lastDot);
	var Octet = address.substr(lastDot, secondColon - lastDot);

	var To = "SupportTeam@Company.com";
	var Subject = "In " + PC;
	var Body = Octet;
	sendMail (To, Subject, Body);
}

function writeLabCallout() {
	var labList = getLablist();
	
	$("#popOutTitle").text("Lab Computer Callout");
	$(".popOutBody").html("\
		<p>\
			Computer : <select id='emailField1'>" + labList + "</select>&nbsp; &nbsp; \
			<input type='radio' name='emailField2' id='emailField2_1' value='In' checked><label for='emailField2_1'>In</label> &nbsp; \
			<input type='radio' name='emailField2' id='emailField2_2' value='Out'><label for='emailField2_2'>Out</label> &nbsp; &nbsp; \
			<label for='emailField3'>Connect? </label><input type='checkbox' id='emailField3'>\
		</p>\
		<button onclick='sendLabCallout()'>C<u>r</u>eate Email</button>"
	);
	$(".popOut").show();
	$("#emailField1").focus();
	
	function getLablist() {
		var list = "";
		$(".labLink").each(function(i,n) {
			list += "<option value='" + n.href + "'>" + n.text + "</option>";	
		});
		return list;
	}
}

function sendLabCallout(lab) {
	var launchLab = $("#emailField3").is(":checked")
	
	if($("#popOutTitle").text() == "Lab Computer Callout") {
		let emailField1 = $("#emailField1 :selected");
		var selectedComputer = emailField1.text();
		var selectedAddress = emailField1.val();
		var callOut = $("input[name='emailField2']:checked").val();
	} else {
		var selectedComputer = $("#" + lab).text();
		var selectedAddress = $("#" + lab).attr("href")
		var callOut = "In"
	}
	
	let lastDot = selectedAddress.lastIndexOf(".");
	let secondColon = selectedAddress.indexOf(":",lastDot);
	var Octet = selectedAddress.substr(lastDot, secondColon - lastDot);
	
	var To = "SupportTeam@Company.com";
	var Subject = callOut + " : " + selectedComputer;
	var Body = Octet;
	sendMail (To, Subject, Body);
	
	if(launchLab) { window.open(selectedAddress) }
}

function writeCallTags() {
	let loadedStore = $('#storeNumber').text();
	
	$('#popOutTitle').text("Call Tags Email");
	$('.popOutBody').html("\
		<p>\
		Store : <input id='emailField1' type='number' value='" + loadedStore + "' maxlength=5 size='5'>\
		&nbsp; &nbsp; Dud #<input id='emailField2' type='number'>\
		</p>\
		<button onclick='sendCallTags()'>C<u>r</u>eate Email</button>\
	");
	$(".popOut").show();
	$("#emailField1").focus();
}

function sendCallTags() {
	let Store = $('#emailField1').val();
	let DUD = $('#emailField2').val();
	
	try {
		if (Store == "") throw "Please enter the store number";
		if (DUD == "") throw "Please enter the DUD to be returnd";
	} catch(erMsg) {
		alert( erMsg );
		return;
	};
	
	let loc = get_store_location(Store);
	
	console.log(loc);
	
	let To = "Support@Vendor.com";
	let Copy = "SupportTeam@Company.com";
	let Subject = "Pickup for DUD " + encodeURIComponent("#") + DUD;
	let Body = "Store " + Store + " in " + loc.city + ", " + loc.state + " has reported that they lost or did not receive a return shipping label for DUD " + encodeURIComponent("#") + DUD + ". Please schedule a pickup for this package.";
	sendMail (To, Subject, Body, Copy);
	
	function get_store_location(store){
		let data = $.grep(storeData, function(line) {
			return (line.Store == store)
			
		});
		console.log(data);
		let loc = {
			city : data[0].City,
			state : data[0].State
		}
		return loc;
	};
}