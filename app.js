const http = require('http');
const fs = require('fs');
const spawn = require('child_process').spawn;
var argv = require('optimist')
  .default('p', 8080)
  .argv;


var wordExeList = [
  'C:\\Program Files\\Microsoft Office\\Office15\\WINWORD.EXE',
  'C:\\Program Files (x86)\\Microsoft Office\\Office15\\WINWORD.EXE'
];

var wordExec = null;

if (argv.w) {
  if (fs.existsSync(argv.w)) {
    wordExec = argv.w;
  }
} else {
  for (var key in wordExeList) {
    if (fs.existsSync(wordExeList[key])) {
      wordExec = wordExeList[key];
      break;
    }
  }
}

if (wordExec != null) {
  createServer(argv.p);
} else {
  console.error("Word executable not found");
}

function createServer(port) {
  var server = http.createServer(handleRequest);

  server.listen(port, function() {
    console.log("Server listening on: http://localhost:%s", argv.p);
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
    return returnError(response, err);
  }

  const converter = spawn(wordExec, ['/mExportToPDFext', '/q', 'doc.docx']);

  converter.on('close', function(code) {
    if (code !== 0) {
      return returnError(response);
    }
    fs.readFile('doc.pdf', function(err, data) {
      if (err) {
        return returnError(response, err);
      }
      response.end(data);
    });
  });
}

function returnError(response, err) {
  response.statusCode = 500;
  response.end('An error occured');
  return console.log(err);
}
