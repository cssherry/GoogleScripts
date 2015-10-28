function JSONtoRSS(json) {
  try {

    var result = UrlFetchApp.fetch(json);

    if (result.getResponseCode() === 200) {

      var articles = Utilities.jsonParse(result.getContentText());

      if (articles) {
        var results = articles.results;
        var len = articles.results.length;

        var rss = "";

        if (len) {

          rss  = '<?xml version="1.0"?><rss version="2.0" xmlns:dc="http://purl.org/dc/elements/1.1/">';
          rss += ' <channel><title>' + articles.meta.method + '</title>';
          rss += ' <link>' + articles.meta.link + '</link>';
          rss += ' <pubDate>' + new Date().toUTCString() + '</pubDate>';

          for (var i=0; i<len; i++) {
            var group = results[i].group.name;
            var urlname = results[i].group.urlname;
            var addedTime = new Date(results[i].created).toUTCString();
            var eventStartTime = new Date(results[i].time).toString();
            var eventPlace;
            if (results[i].venue) {
              eventPlace = results[i].venue.name + ", " +
                               results[i].venue.address_1 + ", " +
                               results[i].venue.city;
            } else {
              eventPlace = "TBD";
            }
            var eventUrl = results[i].event_url;
            var RSVPLimit = results[i].rsvp_limit;
            var eventName  = results[i].name ;
            var eventDescription  = "Time: " + eventStartTime + "/ Place: " + eventPlace + "\n" + results[i].description;

            rss += "<item><title>" + group + ": " + eventName + "</title>";
            rss += " <dc:creator>" + group + "(" + urlname + ") </dc:creator>";
            rss += " <pubDate>" + addedTime + "</pubDate>";
            rss += " <guid isPermaLink='false'>" + results[i].id + "</guid>";
            rss += " <link>" + eventUrl + "</link>";
            rss += " <description>" + escapeXml(eventDescription) + "</description>";
            rss += "</item>";
          }

          rss += "</channel></rss>";

          return rss;
        }
      }
    }
  } catch (e) {
    Logger.log(e.toString());
  }
}

function doGet(e) {
  var feed = e.parameter.url || url;
  var rss = JSONtoRSS (feed);
  return ContentService.createTextOutput(rss)
    .setMimeType(ContentService.MimeType.RSS);
}

function escapeXml(unsafe) {
  return unsafe.replace(/[<>&'"\n\r]/g, function (c) {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\'': return '&apos;';
            case '"': return '&quot;';
            case '\n': return '&#xD;';
            case '\r': return '&#xD;';
        }
    });
}