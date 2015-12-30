var hostweburl;
var appweburl;
var context;
var appContextSite;
var factory;
var web;
var list;
var listitemcollection1;
var listitemcollection2;
var listitemcollection3;
var user;

$(document).ready(function () {

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
    list = web.get_lists().getByTitle("EmpExpenses");

    HideAll();
    $("#Dashboard").show();
    GetUser();
    ExecuteGetApproved();
    ExecuteGetPending();

});


function HideAll() {    
    document.getElementById("Addform").reset();
    $("#btnAddSubmit").show();
    $("#btnUpdate").hide();
    $("form").submit(function () { return false; });
    $("#Dashboard").hide();
    $("#BeforeAdd").hide();
    $("#AfterAdd").hide();
    $("#PendingExpenses").hide();
    GetUser();    
}

function DashboardClick() {
    HideAll();
    ExecuteGetApproved();
    $("#Dashboard").show();
}

function AddExpenseClick() {
    HideAll();
    $("#AddExpense").show();
    $("#BeforeAdd").show();
    $("#AfterAdd").hide();
    FreezAddForm(false);
}
function UpdateExpensesClick() {
    HideAll();
    ExecuteGetPending();
    $("#PendingExpenses").show();
}

function SubmitAdd() {
    ExecuteAdd();
    HideAll();
    $("#AfterAdd").show();
}

function SubmitUpdate() {
    ExecuteUpdate();
    HideAll();
    $("#AfterAdd").show();
}

function ReAdd() {
    HideAll();
    $("#BeforeAdd").show();
    FreezAddForm(false);
}

// Function to freez controls forr VIEW & EDIT.
function FreezAddForm(value) {
    $('#txtExpenseTitle').attr('readonly', value);
    $('#ddlExpenseType').attr('readonly', value);
    $('#txtExpenseDate').attr('readonly', value);
    $('#txtBillNo').attr('readonly', value);
    $('#txtBillAmount').attr('readonly', value);
    $('#ddlBillcurrency').attr('readonly', value);
    $('#txtExpLocation').attr('readonly', value);
    $('#txtMerDetail').attr('readonly', value);
    $('#txtJustification').attr('readonly', value);
    if (value == true)
        $('#btnAddSubmit').hide();
    else
        $('#btnAddSubmit').show();
}

// Function to retrieve a query string value.
function getQueryStringParameter(paramToRetrive) {
    var params = document.URL.split("?")[1].split("&");
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrive) return singleParam[1];
    }
}

//************ Load User data from AD*****************

// funtion to retrive User details from AD.
function GetUser() {
    user = context.get_web().get_currentUser();
    context.load(user);
    context.executeQueryAsync(GetUserSuccess, GetUserFail);
}

function GetUserSuccess(data, req) {
    $('#lblUserName').text(user.get_title());
}

function GetUserFail(data, error, errorMessage) {
    console.log('Error at GetUser() : ' + errorMessage);
}

//************ Add Expenses*****************

function ExecuteAdd() {
    var rate = CurrencyExchange($("#lblBaseCurrency").text(), $("#ddlBillcurrency").val());    
    var reImbAmount = (rate) * ($("#txtBillAmount").val());
    var shortReImbAmount = reImbAmount.toFixed(2);

    var listItemCreationInfo = new SP.ListItemCreationInformation();
    var newItem = list.addItem(listItemCreationInfo);
    newItem.set_item('Title', $("#txtExpenseTitle").val());
    newItem.set_item('ExpenseType', $("#ddlExpenseType").val());
    newItem.set_item('ExpenseDate', $("#txtExpenseDate").val());
    newItem.set_item('BillNo', $("#txtBillNo").val());
    newItem.set_item('BillAmount', $("#txtBillAmount").val());
    newItem.set_item('Billcurrency', $("#ddlBillcurrency").val());
    newItem.set_item('ExpenseLocation', $("#txtExpLocation").val());
    newItem.set_item('MerchantDetails', $("#txtMerDetail").val());
    newItem.set_item('Justification', $("#txtJustification").val());
    newItem.set_item('ConversionRate', rate);
    newItem.set_item('ReimbursementAmount', shortReImbAmount);
    newItem.set_item('EmployeeID', $("#lblEmpId").text());
    newItem.update();
    context.load(newItem);
    context.executeQueryAsync(ExecuteAddSuccess, ExecuteAddError);
}

function ExecuteAddSuccess(data, req) {
    $("#lblAddSubmitMessage").text("Your details has been sucessfully submitted !");
}
function ExecuteAddError(data, error, errorMessage) {
    $("#lblAddSubmitMessage").text("Oops! Something went wrong! Please try again");
    console.log("Error in ExecuteAdd() : " + errorMessage);
}

//************ Load Expenses*****************

function ExecuteGet(id, edit) {
    $("#Expid").text(id);
    AddExpenseClick();
    if (edit == false)
        FreezAddForm(true);
    else {
        $("#btnAddSubmit").hide();
        $("#btnUpdate").show();
    }
    var camlString1 = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + id + "</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>"
    var camlQuery1 = new SP.CamlQuery();
    camlQuery1.set_viewXml(camlString1);
    listitemcollection1 = list.getItems(camlQuery1);
    context.load(listitemcollection1, "Include(ID, Title, ExpenseType, ExpenseDate, BillNo, BillAmount, Billcurrency, ExpenseLocation, MerchantDetails, Justification)");
    context.executeQueryAsync(ExecuteGetSuccess, ExecuteGetError);
}

function ExecuteGetSuccess(data, req) {
    var enumerator = listitemcollection1.getEnumerator();
    while (enumerator.moveNext()) {
        var results = enumerator.get_current();        
        $("#txtExpenseTitle").val(results.get_item("Title"));
        $("#ddlExpenseType").val(results.get_item("ExpenseType"));
        $("#txtExpenseDate").val(results.get_item("ExpenseDate"));
        $("#txtBillNo").val(results.get_item("BillNo"));
        $("#txtBillAmount").val(results.get_item("BillAmount"));
        $("#ddlBillcurrency").val(results.get_item("Billcurrency"));
        $("#txtExpLocation").val(results.get_item("ExpenseLocation"));
        $("#txtMerDetail").val(results.get_item("MerchantDetails"));
        $("#txtJustification").val(results.get_item("Justification"));
    }
}

function ExecuteGetError(data, error, errorMessage) {
    colsole.log("Error in ExecuteGetError(): " + errorMessage);
}


//************ Approved Expenses*****************


function ExecuteGetApproved() {
    var camlString2 = "<View><Query><Where><Eq><FieldRef Name='ApprovalStatus' /><Value Type='Choice'>Approved</Value></Eq></Where></Query></View>";
    var camlQuery2 = new SP.CamlQuery();
    camlQuery2.set_viewXml(camlString2);
    listitemcollection2 = list.getItems(camlQuery2);
    context.load(listitemcollection2, "Include(ID, Title, ExpenseDate, BillAmount, Billcurrency, ReimbursementAmount)");
    context.executeQueryAsync(ExecuteGetApprovedSuccess, ExecuteGetApprovedError);
}

function ExecuteGetApprovedSuccess(data, req) {
    var tblHeader = "<thead><tr class=\"row1\"><th class=\"cell1\">Date</th><th class=\"cell1\">Bill Amount</th><th class=\"cell1\">Amount Reimbursed</th><th>Expense Title</th></tr></thead>";
    var tblBodyRows = "";

    var enumerator = listitemcollection2.getEnumerator();
    while (enumerator.moveNext()) {
        var results = enumerator.get_current();
        tblBodyRows = tblBodyRows + "<tr>";
        tblBodyRows = tblBodyRows + "<td>" + results.get_item("ExpenseDate") + "</td>";
        tblBodyRows = tblBodyRows + "<td>" + results.get_item("BillAmount") + "&nbsp;" + results.get_item("Billcurrency") + "</td>";
        tblBodyRows = tblBodyRows + "<td>" + results.get_item("ReimbursementAmount") + "&nbsp;"+ $("#lblBaseCurrency").text() + "</td>";
        tblBodyRows = tblBodyRows + "<td><a href=\"javascript:ExecuteGet(" + results.get_item("ID") + ", false);\">" + results.get_item("Title") + "</a></td>";
        tblBodyRows = tblBodyRows + "</tr>";
    }
    var tblContent = "<table class=\"tbl\">" + tblHeader + "<tbody>" + tblBodyRows + "</tbody></table>";

    if (tblBodyRows != "")
        $("#ApprovedExpenses").html(tblContent);

    else
        $("#ApprovedExpenses").html("<b>There is no recently approved expenses.</b>");
}

function ExecuteGetApprovedError(data, error, errorMessage) {
    colsole.log("Error in ExecuteGetApproved(): " + errorMessage);
}


//************ Pending Expenses*****************

function ExecuteGetPending() {

    var camlString3 = "<View><Query><Where><Eq><FieldRef Name='ApprovalStatus' /><Value Type='Choice'>Pending</Value></Eq></Where></Query></View>";
    var camlQuery3 = new SP.CamlQuery();
    camlQuery3.set_viewXml(camlString3);
    listitemcollection3 = list.getItems(camlQuery3);
    context.load(listitemcollection3, 'Include(ID, Title, ExpenseDate, BillAmount, Billcurrency, ReimbursementAmount)');
    context.executeQueryAsync(ExecuteGetPendingSuccess, ExecuteGetPendingError);
}

function ExecuteGetPendingSuccess(data, req) {
    var tblHeader = "<thead><tr class=\"row1\"><th class=\"cell1\"></th><th class=\"cell1\">Date</th><th class=\"cell1\">Bill Amount</th><th class=\"cell1\">Amount Reimbursed</th><th>Expense Title</th></tr></thead>";
    var tblBodyRows = "";

    var enumerator = listitemcollection3.getEnumerator();
    while (enumerator.moveNext()) {
        var results = enumerator.get_current();
        tblBodyRows = tblBodyRows + "<tr>";
        tblBodyRows = tblBodyRows + "<td><input type=\"button\" onclick=\"ExecuteGet(" + results.get_item("ID") + ", true); \" value=\"Edit\"></td>";
        tblBodyRows = tblBodyRows + "<td>" + results.get_item("ExpenseDate") + "</td>";
        tblBodyRows = tblBodyRows + "<td>" + results.get_item("BillAmount") + "&nbsp;" + results.get_item("Billcurrency") + "</td>";
        tblBodyRows = tblBodyRows + "<td>" + results.get_item("ReimbursementAmount") + "&nbsp;" + $("#lblBaseCurrency").text() + "</td>";
        tblBodyRows = tblBodyRows + "<td><a href=\"javascript:ExecuteGet(" + results.get_item("ID") + ", false);\">" + results.get_item("Title") + "</a></td>";
        tblBodyRows = tblBodyRows + "</tr>";
    }
    var tblContent = "<table class=\"tbl\">" + tblHeader + "<tbody>" + tblBodyRows + "</tbody></table>";

    if (tblBodyRows != "")
        $("#PendingExpenses").html(tblContent);

    else
        $("#PendingExpenses").html("<b>There is no recently approved expenses.</b>");
}

function ExecuteGetPendingError(data, error, errorMessage) {
    alert("Error in ExecuteGetPending(): " + errorMessage);
}


/************************/
function ExecuteUpdate() {        
        var Eid = $("#Expid").text();
        var oListItem = list.getItemById(Eid);
        oListItem.set_item('Title', $("#txtExpenseTitle").val().toString());
        oListItem.set_item('ExpenseType', $("#ddlExpenseType").val().toString());
        oListItem.set_item('ExpenseDate', $("#txtExpenseDate").val().toString());
        oListItem.set_item('BillNo', $("#txtBillNo").val().toString());
        oListItem.set_item('BillAmount', $("#txtBillAmount").val().toString());
        oListItem.set_item('Billcurrency', $("#ddlBillcurrency").val().toString());
        oListItem.set_item('ExpenseLocation', $("#txtExpLocation").val().toString());
        oListItem.set_item('MerchantDetails', $("#txtMerDetail").val().toString());
        oListItem.set_item('Justification', $("#txtJustification").val().toString());
        oListItem.update();
        context.executeQueryAsync(ExecuteUpdateSuccess, ExecuteUpdateError);    
}

function ExecuteUpdateSuccess(data, req) {    
    $("#lblAddSubmitMessage").text("Your details has been sucessfully updated !");    
}

function ExecuteUpdateError(data, error, errorMessage) {
    $("#lblAddSubmitMessage").text("Oops! Something went wrong! Please try again");
    console.log("Error in ExecuteAdd() : " + errorMessage);
}


function CurrencyExchange(basecurrency, billcurrency) {
    var exchangeRate = {
        "rates": [
        {
            "Currency": "USD",
            "Rate": "0.015731"
        },
        {
            "Currency": "EUR",
            "Rate": "0.013853"
        },
        {
            "Currency": "INR",
            "Rate": "1"
        },
        {
            "Currency": "GBP",
            "Rate": "0.009996"
        },
        {
            "Currency": "AUD",
            "Rate": "0.019622"
        }
        ]
    };

    var rateBillingtoINR;
    var rateINRtoBase;

    var data = exchangeRate.rates;

    for (var i in data) {
        if (data[i].Currency == billcurrency)
            rateBillingtoINR = data[i].Rate;
    }

    for (var i in data) {
        if (data[i].Currency == basecurrency)
            rateINRtoBase = data[i].Rate;
    }

    var finalRate = rateINRtoBase / rateBillingtoINR;
    return finalRate;
}

// Fixer.io
// JSON API for foreign exchange rates and currency conversion

function CurrencyExchangeRate() {
    try {
        var rateBillingtoIndex;
        var rateIndextoBase;
        var finalrate;
        
        basecurrency = $("#lblBaseCurrency").text();
        billcurrency = $("#ddlBillcurrency").val();

        $.getJSON('http://api.fixer.io/latest',
              function (data) {
                  fx.rates = data.rates;
                  var rate = fx(1).from(billcurrency).to(basecurrency);
                  var round = Math.round(rate * 100) / 100;
                  alert(round);
                  finalrate = round;
              }
        );
        alert(finalrate);
        //return finalrate;
    }
    catch (e) {
        alert(e.message);
    }
}