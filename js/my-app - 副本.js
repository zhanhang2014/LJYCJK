// Let's register Template7 helper so we can pass json string in links
Template7.registerHelper('json_stringify', function (context) {
    return JSON.stringify(context);
});

// Initialize your app
var myApp = new Framework7({
    animateNavBackIcon: true,
    // Enable templates auto precompilation
    precompileTemplates: true,
    // Enabled pages rendering using Template7
    template7Pages: true,
    //no cache
    cache: false,
    // Specify Template7 data for pages
    template7Data: {
        // cable crane data
        cable_crane_1_1: {
            
        },
        cable_crane_1_2: {
            
        },
        cable_crane_1_3:{

        },
        cable_crane_1_4:{

        }
    }
});

// Export selectors engine
var $$ = Dom7;

// Add main View
var mainView = myApp.addView('.view-main', {
    // Enable dynamic Navbar
    dynamicNavbar: true,
});

var refreashState = true;
var refreashTimer;

myApp.onPageInit('crane_temp', function (page) {
  //console.log('refreash page');
  //startRefreashData();
  if(refreashState){
    refreashTimer = window.setInterval(getAjaxData,1000);
    refreashState = false;
  }
});

//var refreashTimer;

function startRefreashData(){
    //refreashTimer=window.setInterval(getAjaxData,1000);
    console.log("trigger refresh");
    getAjaxData();
}

function getAjaxData(){
    loadXMLDoc();
    //window.clearInterval(refreashTimer);
    console.log('refreash page');
    mainView.router.load({
    ignoreCache: true,
    reload: true,
    template: Template7.templates.crane_data, 
    context: myApp.template7Data.cable_crane_1_1
  });
}

function stopRefreashData(){
    console.log('stop refreash page');
    window.clearInterval(refreashTimer);
    refreashState = true;
}

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

function loadXMLDoc()
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
    xmlhttp.open("GET","http://localhost:3000/dist/LJ_OPC_WINCC.xlsm",true);
    //xmlhttp.open("GET","http://10.180.84.197:3000/dist/LJ_OPC_WINCC.xlsx",true);

    if(typeof Uint8Array !== 'undefined') {
        xmlhttp.responseType = "arraybuffer";
        xmlhttp.onload = function(e) {
            var arraybuffer = xmlhttp.response;
            var data = new Uint8Array(arraybuffer);
            var arr = new Array();
            for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
            var wb = XLSX.read(arr.join(""), {type:"binary"});
            myApp.template7Data.cable_crane_1_1=process_wb(wb);
            myApp.template7Data.cable_crane_1_2=process_wb(wb);
            myApp.template7Data.cable_crane_1_3=process_wb(wb);
            myApp.template7Data.cable_crane_1_4=process_wb(wb);
        };
    } else {
        xmlhttp.setRequestHeader("Accept-Charset", "x-user-defined");  
        xmlhttp.onreadystatechange = function() { if(xmlhttp.readyState == 4 && xmlhttp.status == 200) {
            var ff = convertResponseBodyToText(xmlhttp.responseBody);
            var wb = XLSX.read(ff, {type:"binary"});
            myApp.template7Data.cable_crane_1_1=process_wb(wb);
            myApp.template7Data.cable_crane_1_2=process_wb(wb);
            myApp.template7Data.cable_crane_1_3=process_wb(wb);
            myApp.template7Data.cable_crane_1_4=process_wb(wb);
        } };
    }
    xmlhttp.send();
}