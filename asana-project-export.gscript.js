// Sets Asana API Key
function AuthData()
{
	return {"Authorization": "Basic " + Utilities.base64Encode("" + ":" + "")};
}

var workspace = "";

// Simulates PHP's date function
Date.prototype.format=function(e){var t="";var n=Date.replaceChars;for(var r=0;r<e.length;r++){var i=e.charAt(r);if(r-1>=0&&e.charAt(r-1)=="\\"){t+=i}else if(n[i]){t+=n[i].call(this)}else if(i!="\\"){t+=i}}return t};Date.replaceChars={shortMonths:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],longMonths:["January","February","March","April","May","June","July","August","September","October","November","December"],shortDays:["Sun","Mon","Tue","Wed","Thu","Fri","Sat"],longDays:["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],d:function(){return(this.getDate()<10?"0":"")+this.getDate()},D:function(){return Date.replaceChars.shortDays[this.getDay()]},j:function(){return this.getDate()},l:function(){return Date.replaceChars.longDays[this.getDay()]},N:function(){return this.getDay()+1},S:function(){return this.getDate()%10==1&&this.getDate()!=11?"st":this.getDate()%10==2&&this.getDate()!=12?"nd":this.getDate()%10==3&&this.getDate()!=13?"rd":"th"},w:function(){return this.getDay()},z:function(){var e=new Date(this.getFullYear(),0,1);return Math.ceil((this-e)/864e5)},W:function(){var e=new Date(this.getFullYear(),0,1);return Math.ceil(((this-e)/864e5+e.getDay()+1)/7)},F:function(){return Date.replaceChars.longMonths[this.getMonth()]},m:function(){return(this.getMonth()<9?"0":"")+(this.getMonth()+1)},M:function(){return Date.replaceChars.shortMonths[this.getMonth()]},n:function(){return this.getMonth()+1},t:function(){var e=new Date;return(new Date(e.getFullYear(),e.getMonth(),0)).getDate()},L:function(){var e=this.getFullYear();return e%400==0||e%100!=0&&e%4==0},o:function(){var e=new Date(this.valueOf());e.setDate(e.getDate()-(this.getDay()+6)%7+3);return e.getFullYear()},Y:function(){return this.getFullYear()},y:function(){return(""+this.getFullYear()).substr(2)},a:function(){return this.getHours()<12?"am":"pm"},A:function(){return this.getHours()<12?"AM":"PM"},B:function(){return Math.floor(((this.getUTCHours()+1)%24+this.getUTCMinutes()/60+this.getUTCSeconds()/3600)*1e3/24)},g:function(){return this.getHours()%12||12},G:function(){return this.getHours()},h:function(){return((this.getHours()%12||12)<10?"0":"")+(this.getHours()%12||12)},H:function(){return(this.getHours()<10?"0":"")+this.getHours()},i:function(){return(this.getMinutes()<10?"0":"")+this.getMinutes()},s:function(){return(this.getSeconds()<10?"0":"")+this.getSeconds()},u:function(){var e=this.getMilliseconds();return(e<10?"00":e<100?"0":"")+e},e:function(){return"Not Yet Supported"},I:function(){var e=null;for(var t=0;t<12;++t){var n=new Date(this.getFullYear(),t,1);var r=n.getTimezoneOffset();if(e===null)e=r;else if(r<e){e=r;break}else if(r>e)break}return this.getTimezoneOffset()==e|0},O:function(){return(-this.getTimezoneOffset()<0?"-":"+")+(Math.abs(this.getTimezoneOffset()/60)<10?"0":"")+Math.abs(this.getTimezoneOffset()/60)+"00"},P:function(){return(-this.getTimezoneOffset()<0?"-":"+")+(Math.abs(this.getTimezoneOffset()/60)<10?"0":"")+Math.abs(this.getTimezoneOffset()/60)+":00"},T:function(){var e=this.getMonth();this.setMonth(0);var t=this.toTimeString().replace(/^.+ \(?([^\)]+)\)?$/,"$1");this.setMonth(e);return t},Z:function(){return-this.getTimezoneOffset()*60},c:function(){return this.format("Y-m-d\\TH:i:sP")},r:function(){return this.toString()},U:function(){return this.getTime()/1e3}}


function GetProjects()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Live source data');

  var projects = [];
  var offset = "";
  var options = { "method" : "get"};
  options.headers = AuthData();
  options.headers["Asana-Fast-Api"] = "true";

  var x = 0;
  do
  {
    var urlpath = "https://app.asana.com/api/1.0/projects?limit=100&workspace=" + workspace + "&archived=false";
    if (offset != null && offset !== '')
      urlpath += "&offset=" + offset;

    var response = UrlFetchApp.fetch(urlpath, options);

    jResponse = JSON.parse (response);
    if (jResponse.next_page != null && jResponse.next_page.offset != null)
      offset = jResponse.next_page.offset;
    else
      offset = '';

    var data = jResponse.data;

    for(var i in data)
    {
      projects.push(data[i].id);
    }
    x++;
  } while (offset != '');

  Logger.log("Projects:\n" + projects);

  //throw new Error("Got all the projects! :D");

  var rangesToClear = ['A3:K1000'];
  for (var i=0; i<rangesToClear.length; i++) {
	   sheet.getRange(rangesToClear[i]).clearContent();
  }

  var currentRow=3; // Skip title rows
  for(var i in projects)
  {
    var options = { "method" : "get"};
    options.headers = AuthData();
    var projectResponse = UrlFetchApp.fetch("https://app.asana.com/api/1.0/projects/"+projects[i], options);

    projectData = JSON.parse (projectResponse);
    projectData = projectData.data;

	// Skip logging archived projects
	if (projectData.archived)
		continue;

	// Basic housekeeping for some of the data to prevent errors
    if (!projectData.owner || !projectData.owner.name)
        projectData.owner = {"name": "None"};
    if (!projectData.notes)
        projectData.notes = "";

	// Check if a status has ever been set, skip fields if not
    if (projectData.current_status)
		rowArray = [ projectData.workspace.name, projectData.name, "https://app.asana.com/0/"+projectData.id, projectData.owner.name, projectData.archived, projectData.team.name, projectData.notes, projectData.current_status.color, projectData.current_status.text, projectData.current_status.author.name, Utilities.formatDate(new Date(projectData.current_status.modified_at), "GMT-5", "MM/dd/yyyy") ];
	else
		rowArray = [ projectData.workspace.name, projectData.name, "https://app.asana.com/0/"+projectData.id, projectData.owner.name, projectData.archived, projectData.team.name, projectData.notes, "", "", "", "" ];

	// Format inside object for the setValues function
	var newData = [rowArray];
	sheet.getRange(currentRow,1,1,rowArray.length).setValues(newData);

	currentRow++;

  }
}
