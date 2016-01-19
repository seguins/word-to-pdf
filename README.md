# Api Converter Word to PDF

You need Windows and Microsoft Word

## Install

1. Install the macro script : https://github.com/oleksiykovtun/Word-Export-to-PDF
2. ```npm install```

## Usage

Start server :
 * ```node app.js```
 * ```node app.js -p 8080```
 * ```node app.js -w /path/word.exe```

Exemple request :
 * ```curl --data-binary @test.docx 127.0.0.1:8080 > test.pdf```
