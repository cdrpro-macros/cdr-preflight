// JavaScript Document
var app = window.external.Application;

function OnLoad() {
	app.InitializeVBA();
	RegisterEvent();
	app.GMSManager.RunMacro("CdrPreflight", "cp_Info.InitializeCdrPreflight");
	list2.innerHTML = "";
	cmLoadPresets();
	cmLoadConvPresets();
	cmRefresh();
}

function RegisterEvent() {
	window.external.RegisterEventListener( "CloseDocument", "OnCloseDocument()" );
}
function UnregisterEvent() {
	window.external.UnregisterEventListener( "CloseDocument" );
}

function cmOpenRegWindow() {
	app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.cmOpenRegWindow");
}

function cmRefresh() {
	var list = app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.cmRefresh", presetsListSel.value);
	ListItems.innerHTML = list;
	list2.innerHTML = "";
	cmLoadPresets();
	cmLoadConvPresets();
}
function OnCloseDocument() {
	ListItems.innerHTML = '<p style="background:none; padding-left:4px; color:#F60;">Click on the <b>Refresh</b> button for get information about active document.</p>';
	list2.innerHTML = '';
}

function cmOptions() {
	app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.cmOptions");
}


function cmHelp() {
	app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.ShowHelp");
}
function cmAbout() {
	var list = app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.ShowAbout");
	ListItems.innerHTML = list;
}


function cmLoadPresets() {
	var list = app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.cmLoadPresets");
	presetsList.innerHTML = list;
}
function cmChangePreset() {
	app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.cmChangePreset", presetsListSel.value);
	cmRefresh()
}


function ExportList() {
	app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.ExportList", ListItems.innerHTML);
}


function LoadList2(typeID, label) {
	var list = app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.LoadList2", typeID);
	list2.innerHTML = list;
	activeTYPE.innerHTML = label + " (" + app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.CountList2") + ")";
}
function SelectShape(item) {
	app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.SelectShapeByItem", item);
}


function cmLoadConvPresets() {
	var list = app.GMSManager.RunMacro("CdrPreflight", "cp_Convert.cmLoadConvPresets");
	ConvPresets.innerHTML = list;
}
function cmChangeConvPreset() {
	app.GMSManager.RunMacro("CdrPreflight", "cp_Convert.cmChangeConvPreset", ConvPresetsListSel.value);
}
function cmConverter() {
	var list = app.GMSManager.RunMacro("CdrPreflight", "cp_Convert.cmConverter", ConvPresetsListSel.value);
	cmRefresh();
	list2.innerHTML = list;
}
function cmConvOpt() {
	app.GMSManager.RunMacro("CdrPreflight", "ctc_docker.cmConvOpt");
}