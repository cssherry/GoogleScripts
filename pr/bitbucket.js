// Pull PR info from bitbucket and merge into spreadsheet
function syncPRs() {

  var releases = [ 'POC', 'Crater Lake', 'Canyonlands', 'Denali', 'Everglades', 'Grand Canyon', '5.0.1' ];
  var statuses = [ 'Ready', 'Not Ready', 'WIP', 'Work in Progress', 'Not to Merge', 'BLOCKER', 'Critical' ];

  var headers = personalAuth;
  var response = UrlFetchApp.fetch(bitbucketLink, { headers: headers, contentType: "application/json"});
  var data = JSON.parse(response.getContentText());
  var openprs = data.values;
  var response = UrlFetchApp.fetch("https://api.bitbucket.org/2.0/repositories/arcadiadata/arcviz/pullrequests?pagelen=50&state=MERGED", { headers: headers, contentType: "application/json"});
  var data = JSON.parse(response.getContentText());
  var mergedprs = data.values;
  var response = UrlFetchApp.fetch("https://api.bitbucket.org/2.0/repositories/arcadiadata/arcviz/pullrequests?pagelen=50&state=DECLINED", { headers: headers, contentType: "application/json"});
  var data = JSON.parse(response.getContentText());
  var declinedprs = data.values;
  var allprs = [];
  allprs = allprs.concat(openprs, mergedprs, declinedprs);

  var coredata = Object.keys(allprs).map(function(k) { var d = allprs[k]; return { id: d.id, title: d.title, author: d.author.display_name, created_on: d.created_on, updated_on: d.updated_on, state: d.state, url: d.links.html.href};})
//  Logger.log(response.getContentText());
  // Logger.log(Object.keys(data.values));
//  Logger.log(coredata);

  var datamap = {};
  for (var i in coredata) {
    var d = coredata[i]
    var title = d.title;

    for (var r in releases) {
      var release = releases[r];
      if (title.toLowerCase().indexOf('[' + release.toLowerCase() + ']') >= 0) {
        d.release = release;
        break;
      } else {
        d.release = "";
      }
    }

    for (var s in statuses) {
      var status = statuses[s];
      if (title.toLowerCase().indexOf('[' + status.toLowerCase() + ']') >= 0) {
        d.status = status;
        break;
      } else {
        d.status = "";
      }
    }

    // ' + bitLink + '/issues/?jql=id%20in%20(ARC-9577%2C%20ARC-10817%2C%20ARC-5863)

    var jiras = title.match(/(?:ARC|arc|Arc)-\d+/gm);
    if (jiras) {
      if (jiras.length === 1) {
        d.jiras = '=HYPERLINK("' + bitLink + '/browse/' + jiras[0] + '", "' + jiras[0] +'")';
      } else {
        d.jiras = '=HYPERLINK("' + bitLink + '/issues/?jql=id%20in%20(' + jiras.join("%2C%20") + ')", "' + jiras.join(", ") +'")';
        //d.jiras = jiras.join(", ");
      }
    } else {
      d.jiras = "";
    }



    d.id_link = '=HYPERLINK("' + d.url + '", "' + d.id +'")';

    d.created_on = new Date(d.created_on.match(/\d+-\d+-\d+T\d+:\d+:\d+\.\d{3}/gm)[0] + 'Z');
    d.updated_on = new Date(d.updated_on.match(/\d+-\d+-\d+T\d+:\d+:\d+\.\d{3}/gm)[0] + 'Z');

    // d.age = Math.round((new Date() - d.created_on) / (24 * 60 * 60 * 1000));
    // d.staleness = Math.round((new Date() - d.updated_on) / (24 * 60 * 60 * 1000));
    d.age = '=DATEDIF(DATEVALUE("' + (d.created_on.toISOString().replace("T", " ").replace("Z",'')) + '"), NOW(), "D")';
    d.staleness = '=DATEDIF(DATEVALUE("' + (d.updated_on.toISOString().replace("T", " ").replace("Z",'')) + '"), NOW(), "D")';

    datamap[String(d.id)] = [d.id_link, d.title, d.release, d.status, d.jiras, d.author, d.state, d.age, d.staleness];

  }


 // var output = coredata.map(function(d) {return [d.id_link, d.release, d.status, d.title, d.author, d.state, d.age, d.staleness];}) ;
 // Logger.log(output);


  // Open the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Get the current data
  var current_data = sheet.getDataRange().getValues();

  // Update the current PR data based on id
  var output = []
  for (var i=1; i<current_data.length; i++ ) {
    var id = String(Math.floor(current_data[i][0]));
    if (datamap.hasOwnProperty(id)) {
      output.push(datamap[id]);
      delete datamap[id];
    } else {
      output.push(current_data[i].slice(0,9));
    }
  }

  // Add any new PRs
  for (var p in datamap) {
    if (datamap[p][6] === 'OPEN') {
      output.push(datamap[p]);
    }
  }

  // Write the updated data to the sheet
  var len = output.length
  sheet.getRange(2,1,len,9).setValues(output);

  // Sort the Sheet by Activity
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort(9)

}


// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Bitbucket')
      .addItem('Sync PRs','syncPRs')
      .addToUi();
}
