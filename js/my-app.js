// Let's register Template7 helper so we can pass json string in links
Template7.registerHelper('json_stringify', function (context) {
    return JSON.stringify(context);
});

var xlsFile = {};
// Initialize your app
var myApp = new Framework7({
    //Disable App's automatica initialization
    init: false, 
    // Enable templates auto precompilation
    animateNavBackIcon: true,
    // Enabled pages rendering using Template7
    precompileTemplates: true,
    template7Pages: true,
    //no cache
    cache: false,
    // Specify Template7 data for pages
    template7Data: {
        // cable crane data
        cable_crane_1_1:{
            Sheet1:[{}]
        },
        cable_crane_1_2:{
            Sheet1:[{}]
        },
        cable_crane_1_3:{
            Sheet1:[{}]
        },
        cable_crane_1_4:{
            Sheet1:[{}]
        },
        cable_crane_2_1:{
            Sheet1:[{}]
        },
        cable_crane_2_2:{
            Sheet1:[{}]
        },
        cable_crane_2_3:{
            Sheet1:[{}]
        },
        cable_crane_2_4:{
            Sheet1:[{}]
        },
        cable_crane_3_1:{
            Sheet1:[{}]
        },
        cable_crane_3_2:{
            Sheet1:[{}]
        },
        cable_crane_3_3:{
            Sheet1:[{}]
        },
        cable_crane_3_4:{
            Sheet1:[{}]
        },
        cable_crane_4_1:{
            Sheet1:[{}]
        },
        cable_crane_4_2:{
            Sheet1:[{}]
        },
        cable_crane_4_3:{
            Sheet1:[{}]
        },
        cable_crane_4_4:{
            Sheet1:[{}]
        },
        cable_crane_5_1:{
            Sheet1:[{}]
        },
        cable_crane_5_2:{
            Sheet1:[{}]
        },
        cable_crane_5_3:{
            Sheet1:[{}]
        },
        cable_crane_5_4:{
            Sheet1:[{}]
        },
        cable_crane_6_1:{
            Sheet1:[{}]
        },
        cable_crane_6_2:{
            Sheet1:[{}]
        },
        cable_crane_6_3:{
            Sheet1:[{}]
        },
        cable_crane_6_4:{
            Sheet1:[{}]
        },
        site_names:{

        }
    }
});

// when the home page load
myApp.onPageInit('index-1', function (page) {
  //console.log('refreash page');
  //startRefreashData();
  console.log("Initialize main page and data");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_1_1");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_2_1");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_2_2");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_3_1");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_3_2");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_4_1");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_4_2");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_5_1");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_5_2");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_6_1");
  loadXMLDoc("LJ0_OPC_WINCC","cable_crane_6_2");

  document.getElementById("site1").innerHTML="工地一:加载中";
  document.getElementById("site2").innerHTML="工地二:加载中";
  document.getElementById("site3").innerHTML="工地三:加载中";
  document.getElementById("site4").innerHTML="工地四:加载中";
  document.getElementById("site5").innerHTML="工地五:加载中";
  document.getElementById("site6").innerHTML="工地六:加载中";

  setTimeout(showName,3000)

});

//maniually initialize home page to activate the function above
myApp.init();


// Export selectors engine
var $$ = Dom7;

// Add main View
var mainView = myApp.addView('.view-main', {
    // Enable dynamic Navbar
    dynamicNavbar: true,
});

var cur_datafile = "LJ0_OPC_WINCC";
var cur_crane = "cable_crane_1_1";
var refreashState = true;
var refreashTimer;



function specifyCrane(datafile,crane){
    cur_datafile = datafile;
    cur_crane = crane;
}

function handleCraneData(crane){
    cur_datafile = datafile;
    cur_crane = crane;
}

myApp.onPageInit('crane_display', function (page) {
  console.log('start refreshing page');
  if(refreashState){
    refreashTimer = window.setInterval(updateCraneData,1000);
    refreashState = false;
  }
});

function showName(){
  console.log("start showing name")
  document.getElementById("site1").innerHTML="工地一:" + myApp.template7Data.site_names.SITE1;
  document.getElementById("site2").innerHTML="工地二:" + myApp.template7Data.site_names.SITE2;
  document.getElementById("site3").innerHTML="工地三:" + myApp.template7Data.site_names.SITE3;
  document.getElementById("site4").innerHTML="工地四:" + myApp.template7Data.site_names.SITE4;
  document.getElementById("site5").innerHTML="工地五:" + myApp.template7Data.site_names.SITE5;
  document.getElementById("site6").innerHTML="工地六:" + myApp.template7Data.site_names.SITE6;
}

function updateCraneData(){
    console.log("updating "+cur_crane+" using "+cur_datafile+" data");
    var strContent = $$(".stay_to_refresh");
    loadXMLDoc(cur_datafile,cur_crane);
    newCraneDisplay = Template7.templates.crane_data(myApp.template7Data[cur_crane]);
    strContent.find('ul').html(newCraneDisplay);
}

function stopRefreashData(){
    console.log('stop refreash page');
    window.clearInterval(refreashTimer);
    refreashState = true;
}

//read excel
function process_wb(workbook) {
    var result={};
    workbook.SheetNames.forEach(function(sheetName) {
        var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if(roa.length > 0){
            result[sheetName] = roa;
        }
    });
    if(typeof console !== 'undefined') console.log("output", new Date());
    return result;
}

//request data using ajax
function loadXMLDoc(datafile,crane)
{
    var xmlhttp;
    if (window.XMLHttpRequest)
    {// code for IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp=new XMLHttpRequest();
    }
    else
    {// code for IE6, IE5
        xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
    }
    xmlhttp.onreadystatechange=function()
    {
        if (xmlhttp.readyState==4 && xmlhttp.status==200)
        {
        
        }
    }
    //xmlhttp.open("GET","http://localhost:3000/dist/"+datafile+".xlsm",true);
    xmlhttp.open("GET","http://117.149.16.29:3000/dist/"+datafile+".xlsm",true);

    if(typeof Uint8Array !== 'undefined') {
        xmlhttp.responseType = "arraybuffer";
        xmlhttp.onload = function(e) {
            var arraybuffer = xmlhttp.response;
            var data = new Uint8Array(arraybuffer);
            var arr = new Array();
            for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
            var wb = XLSX.read(arr.join(""), {type:"binary"});
            //myApp.template7Data[crane]=process_wb(wb);
            xlsFile = process_wb(wb);
            siteName = crane.split("_")[2];
            var craneNum = (parseInt(crane.split("_")[3])-1)*14;
            for( var i=0;i<14;i++){
                myApp.template7Data[crane].Sheet1[i] = {"NAME":xlsFile.Sheet1[i+craneNum]["NAME"+siteName],"DATA":xlsFile.Sheet1[i+craneNum]["DATA"+siteName]};
            }
            myApp.template7Data.site_names = {"SITE1":xlsFile.Sheet1[0].SITE1,"SITE2":xlsFile.Sheet1[0].SITE2,"SITE3":xlsFile.Sheet1[0].SITE3,"SITE4":xlsFile.Sheet1[0].SITE4,"SITE5":xlsFile.Sheet1[0].SITE5,"SITE6":xlsFile.Sheet1[0].SITE6,}
        };
    } else {
        xmlhttp.setRequestHeader("Accept-Charset", "x-user-defined");  
        xmlhttp.onreadystatechange = function() { if(xmlhttp.readyState == 4 && xmlhttp.status == 200) {
            var ff = convertResponseBodyToText(xmlhttp.responseBody);
            var wb = XLSX.read(ff, {type:"binary"});
            //myApp.template7Data[crane]=process_wb(wb);
            xlsFile = process_wb(wb);
            siteName = crane.split("_")[2];
            var craneNum = (parseInt(crane.split("_")[3])-1)*14;
            for( var i=0;i<14;i++){
                myApp.template7Data[crane].Sheet1[i] = {"NAME":xlsFile.Sheet1[i+craneNum]["NAME"+siteName],"DATA":xlsFile.Sheet1[i+craneNum]["DATA"+siteName]};
            }
            myApp.template7Data.site_names = {"SITE1":xlsFile.Sheet1[0].SITE1,"SITE2":xlsFile.Sheet1[0].SITE2,"SITE3":xlsFile.Sheet1[0].SITE3,"SITE4":xlsFile.Sheet1[0].SITE4,"SITE5":xlsFile.Sheet1[0].SITE5,"SITE6":xlsFile.Sheet1[0].SITE6,}
        } };
    }
    xmlhttp.send();
}