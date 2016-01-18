const http = require('http');
const fs = require('fs');
const spawn = require('child_process').spawn;

const PORT = 8080;

createServer(PORT);

function createServer(port) {
  var server = http.createServer(handleRequest);

  server.listen(port, function() {
    console.log("Server listening on: http://localhost:%s", PORT);
  });
}

function handleRequest(request, response) {
  var docBinary = '';
  request.setEncoding('binary')

  request.on('data', function(chunk) {
    docBinary += chunk;
  });

  request.on('end', function() {
    fs.writeFile('doc.docx', docBinary, 'binary', function(err) {
      onSavedInput(response, err)
    });
  });
}

function onSavedInput(response, err) {
  if (err) {
    return returnError(response);
  }

  const converter = spawn('C:\\Program Files (x86)\\Microsoft Office\\Office15\\WINWORD.EXE', ['/mExportToPDFext', '/q', 'doc.docx']);

  converter.on('close', function(code) {
    if (code !== 0) {
      return returnError(response);
    }
    fs.readFile('doc.pdf', function(err, data) {
      if (err) {
        return returnError(response);
      }
      response.end(data);
    });
  });
}

function returnError(response) {
  response.statusCode = 500;
  response.end('An error occured');
  return console.log(err);
}
