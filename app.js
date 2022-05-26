var fs = require('fs');
var XLSX = require("xlsx");
let http = require('http');
let port = 8010;
const excelToJson = require('convert-excel-to-json');


let sheet_data = excelToJson({
    sourceFile: './Calendar.xlsx'
});
var table_output = "";
console.log("hi")
function return_table() {

    table_output = '<table class="table table-striped table-bordered">';

    for (row in sheet_data.Sheet1) {
        table_output += '<tr>';

        for (col in sheet_data.Sheet1[row]) {
            if (row == 0) {

                table_output += '<th>' + `${sheet_data.Sheet1[row][col]}` + '</th>';

            }
            else {

                table_output += '<td>' + `${sheet_data.Sheet1[row][col]}` + '</td>';

            }

        }

        table_output += '</tr>';

    }


    table_output += '</table>';
    table_output += '<script> setInterval(function myTimer(){location.reload();console.log("interval script works")}, 3000);</script>';

}
return_table();
// console.log("table_output", table_output);

//______________

fs.watch("./Calendar.XLSX", (eventType, filename) => {
    console.log("The file ", filename, " was modified!");
    // We can look for different types of changes on a file
    // using the event type like: rename, change, etc.
    console.log("It was a ", eventType, " event type.");
    sheet_data = excelToJson({
        sourceFile: './Calendar.xlsx'
    });
    return_table();

});


const server = http.createServer((req, res) => {
    res.statusCode = 200;
    res.setHeader('Content-Type', 'text/html');
    res.end(`${table_output}`);
});

server.listen(port, () => {
    console.log(`Server is running on port number::${port}`);
});
