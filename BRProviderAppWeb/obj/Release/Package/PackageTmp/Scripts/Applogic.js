// Add-Edit-Update list data from cross domain using JSOM
var hostweburl;
var appweburl = '';
var context;
var appContextSite;
var factory;
var web;
var list;
var billNo;
var billTitle;
var billStatus;
var billId;
var listitemcollection;

//************************************************************************************************
// DOCUMENT READY CODE BLOCK

// Load the required SharePoint libraries
$(document).ready(function () {
    try {
        hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
        appweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
        
        if (appweburl == 'undefined')
            appweburl = localStorage.getItem("AppUrl");
        else
            localStorage.setItem("AppUrl", appweburl);
       
        context = new SP.ClientContext(appweburl);
        factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
        context.set_webRequestExecutorFactory(factory);
        appContextSite = new SP.AppContextSite(context, hostweburl);

        web = appContextSite.get_web();
        list = web.get_lists().getByTitle("BillReimbursement");

        $("#getDetailsInputs").hide();
        $("#newBill").hide();
        $("#changeStatus").hide();

        // Triggering button click events
        $("#btnExe").click(function () {
            billNo = $("#ipBillNo").val();
            if (billNo.length == 0)
                alert('Fields cannot be left empty. Try Again!');
            else
                GetBillData();
        });

        $("#btnBillEntry").click(function () {
            billNo = $("#tbBillNo").val();
            billTitle = $("#tbBillTitle").val();
            billStatus = "In Process";
            if (billNo.length == 0 || billTitle.length == 0)
                alert('Fields cannot be left empty. Try Again!');
            else
                AddBillData();
        });

    } catch (e) {
        alert('Error encountered : ' + e.message);
    }
});

//************************************************************************************************
// CODE BLOCK TO ADD NEW LIST DATA

// Function to prepare and issue the request to add new List data
function AddBillData() {
    var listItemCreationInfo = new SP.ListItemCreationInformation();
    var newItem = list.addItem(listItemCreationInfo);
    newItem.set_item('Title', billTitle);
    newItem.set_item('BillNo', billNo);
    newItem.set_item('Status', billStatus);
    newItem.update();
    context.load(newItem);
    context.executeQueryAsync(AddBillSuccess, AddBillError);
}

// Function to handle the success event for AddBillData.
function AddBillSuccess(data, req) {
    alert("Details added successfully");
    $("#tbBillNo").val('');
    $("#tbBillTitle").val('');
}

// Function to handle the error event for AddBillData
function AddBillError(data, error, errorMessage) {
    alert("Could not complete cross-domain call: " + errorMessage);
}

//************************************************************************************************
// CODE BLOCK TO RETRIVE LIST DATA

// Function to prepare and issue the request to get existing List data
function GetBillData() {
    var camlString = "<View><Query><Where><Eq><FieldRef Name='BillNo' /><Value Type='Text'>" + billNo + "</Value></Eq></Where></Query></View>";
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(camlString);
    listitemcollection = list.getItems(camlQuery);
    context.load(listitemcollection, "Include(Title, BillNo, Status, ID)");
    context.executeQueryAsync(GetBillSuccess, GetBillError);
}

// Function to handle the success event for GetBillData.
function GetBillSuccess(data, req) {
    var innerData = "";
    var enumerator = listitemcollection.getEnumerator();
    while (enumerator.moveNext()) {
        var results = enumerator.get_current();
        innerData = innerData + "<div><span><b>Bill No. : </b></span><span>" + results.get_item("BillNo") + "</span></div><div><span><b>Title : </b></span><span>" + results.get_item("Title") + "</span></div><div><span><b>Status : </b></span><span>" + results.get_item("Status") + "</span></div>";
        billId = results.get_item("ID");
    }

    if (innerData == "")
        innerData = "<span><b>No Results found !</b></span>";
    else {
        if ($("#taskOption").val() == "edit") {
            PopulateStatus();
            $("#changeStatus").show();
        }
    }
    $("#getDetails").show();
    document.getElementById("getDetails").innerHTML = innerData;

}

// Function to handle the error event for GetBillData
function GetBillError(data, error, errorMessage) {
    alert("Error: " + errorMessage);
}

//************************************************************************************************
// CODE BLOCK TO EDIT LIST DATA

// Function to edit List data
function EditBillData() {
    var oListItem = list.getItemById(billId);
    var newStatus = $("#selectStatus").val();
    oListItem.set_item('Status', newStatus);
    oListItem.update();
    context.executeQueryAsync(EditBillSuccess, EditBillError);
}

// Function to handle the success event for EditBillData.
function EditBillSuccess(data, req) {
    $("#btnExe").click();
    alert("Details updated successfully");

}

// Function to handle the error event for EditBillData
function EditBillError(data, error, errorMessage) {
    alert("Error: " + errorMessage);
}


//************************************************************************************************
// CODE BLOCK TO GET DATA FROM QUERY STRING

// Function to retrieve a query string value.
function getQueryStringParameter(paramToRetrive) {
    var params = document.URL.split("?")[1].split("&");
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrive) return singleParam[1];
    }
}


//************************************************************************************************
// CODE BLOCK TO POPULATE LIST DATA TO DROPDOWN MENU 

// Load STATUS of all possible value of choice field of SP List to dropdown list
function PopulateStatus() {
    var field = list.get_fields().getByInternalNameOrTitle('Status');
    listitemcollection = context.castTo(field, SP.FieldChoice);
    context.load(field);
    context.executeQueryAsync(PopulateStatusSuccess, PopulateStatusError);
}

// Function to handle the success event for PopulateStatus.
function PopulateStatusSuccess(data, req) {
    var distinctChoices = listitemcollection.get_choices();
    var selectItemBox = document.getElementById("selectStatus");

    //To clear existing options in dropdown
    if (selectItemBox.hasChildNodes()) {
        while (selectItemBox.childNodes.length >= 1) {
            selectItemBox.removeChild(selectItemBox.firstChild);
        }
    }

    // To add coices to dropdown list.
    var selectOption = document.createElement("option");
    selectOption.value = '--Select--';
    selectOption.innerHTML = '--Select--';
    selectOption.selected = true;
    selectOption.disabled = true;
    selectItemBox.appendChild(selectOption);

    for (var i = 0; i < distinctChoices.length; i++) {
        selectOption = document.createElement("option");
        selectOption.value = distinctChoices[i];
        selectOption.innerHTML = distinctChoices[i];
        selectItemBox.appendChild(selectOption);
    }
}

// Function to handle the error event for PopulateStatus
function PopulateStatusError(data, error, errorMessage) {
    alert("Error: " + errorMessage);
}


// CURRENTLY BELOW IS NOT IN USE -- JUST FOR FUTURE REFERENCE

// Load particular field values of all entered data to dropdown list
function PopulateDataStatus() {
    var camlString = "<View><ViewFields><FieldRef Name='Status' /></ViewFields></View>"
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(camlString);
    listitemcollection = list.getItems(camlQuery);
    context.load(listitemcollection, "Include(Status)");
    context.executeQueryAsync(PopulateDataStatusSuccess, PopulateDataStatusError);
}

// Function to handle the success event for PopulateStatus.
function PopulateDataStatusSuccess(data, req) {
    var listItemEnumerator = listitemcollection.getEnumerator();
    var selectItemBox = document.getElementById("selectStatus");

    //To clear existing options in dropdown except default selection.
    if (selectItemBox.hasChildNodes()) {
        while (selectItemBox.childNodes.length >= 1) {
            selectItemBox.removeChild(selectItemBox.firstChild);
        }
    }

    // To add coices to dropdown list.
    while (listItemEnumerator.moveNext()) {
        var selectOption = document.createElement("option");
        selectOption.value = listItemEnumerator.get_current().get_item('Status');
        selectOption.innerHTML = listItemEnumerator.get_current().get_item('Status');
        selectItemBox.appendChild(selectOption);
    }
}

// Function to handle the error event for PopulateStatus
function PopulateDataStatusError(data, error, errorMessage) {
    alert("Error: " + errorMessage);
}


//************************************************************************************************
//CODE BLOCK TO TRIGGER ACTION ON CHANGE IN TASK SELECTION IN DROPDOWN LIST

// Function to select task
function SelectingOption() {
    $("#newBill").hide();
    $("#changeStatus").hide();
    $("#getDetails").hide();
    $("#getDetailsInputs").hide();
    $("#newBill").hide();

    if ($("#taskOption").val() == "info") {
        $("#getDetailsInputs").show();
    }
    else if ($("#taskOption").val() == "new") {
        $("#newBill").show();
    }
    else if ($("#taskOption").val() == "edit") {
        $("#getDetailsInputs").show();
    }
    else {
        alert('This functionality is currently unavailable');
    }
}
