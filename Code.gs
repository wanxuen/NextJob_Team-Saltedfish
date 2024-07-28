function doGet(e) {
  var page = e.parameter.page || 'Index'; // Default to 'Index' if no page parameter
  console.log(page);
  return HtmlService.createHtmlOutputFromFile(page);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

