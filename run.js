const ExcelJS = require('exceljs');
const https = require('https');

(async function start() {

    console.clear();
    console.log("==========");
    let data = [];

    const workbook = new ExcelJS.Workbook();

    
    //Read file from internet

    //NOTE: If you need to read a file from a local source, you can use instead:
    // let document = await workbook.xlsx.load(./path/to/file.xlsx).catch(ex => { return null });
    // and comment from line 19 to 27

    let file = await get_xlsx("https://url-to-document.com/document.xlsx");

    if (!file.success) {
        console.log(`ERROR: ${file.message}`);
        return;
    }

    //Loads buffer data into exceljs library 
    let document = await workbook.xlsx.load(file.buffer).catch(ex => { return null });

    if (!document) {
        console.log(`ERROR: document => ${document}`);
        return;
    }
    document = JSON.parse(JSON.stringify(document.model))

    //Get all rows from document Rescata todas las rows del documento
    let rows = document.sheets.map(row => { return row.rows })[0];

    //Create an array of keys from the first row of the document.
    let keys = rows[0].cells.map(cell => { return cell.value });

    //Removes first element from array (wich contains the columns name)
    rows.shift();

    //Cleans the array and returns only the value from the cells.
    rows = rows.map(row => { return row.cells.map(cell => { return cell.value }) })

    rows.forEach((row) => {

        let obj = {}
        row.forEach((field, field_index) => {
            if (keys[field_index] && field) {
                let key = keys[field_index];
                obj[key] = field
            }
        })
        data.push(obj);
    })

    //Deletes empty objects from array.
    data = data.filter(element => Object.keys(element).length !== 0)

    //Prints on console the final results.
    console.log(data);

})();


async function get_xlsx(url) {
    return new Promise((resolve, reject) => {

        let data_stream = [];

        //Reads from internet an excel document.
        https.get(url, (response) => {
            //Create data stream from file
            response.on('data', (chunk) => {
                data_stream.push(chunk);
            }).on('end', () => {
                //Craete buffer from data stream 
                let buffer = Buffer.concat(data_stream);
                //Returns Buffer
                return resolve({ success: true, buffer: buffer });
            }).on('error', (err) => {
                return reject({ success: false, message: err })
            })
        });
    })
}

