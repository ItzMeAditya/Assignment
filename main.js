const ExcelJS = require('exceljs');

const wb = new ExcelJS.Workbook();

const fileName = 'Zecko.xlsx';

wb.xlsx.readFile(fileName).then(() => {
    
    const ws1 = wb.getWorksheet('Pre Program Run');
    const ws2 = wb.addWorksheet('Post Program Run');

    const c1 = ws1.getColumn(1);
    
    const headers = [
        {header: 'Website',key: 'web',width : 30},
        {header: 'Categories',key:'ctg',width : 20}
    ]

    ws2.columns = headers;

    c1.eachCell(c => {
        if (c.value.text == 'https://www.headphonezone.in/'){
            ws2.addRow(['https://www.headphonezone.in/','SHOPIFY'])
        }
        else if (c.value.text == 'https://www.boat-lifestyle.com/'){
            ws2.addRow(['https://www.boat-lifestyle.com/','SHOPIFY'])
        }
        else if (c.value.text == 'somemadeupwebsite.com'){
            ws2.addRow(['somemadeupwebsite.com','NOT_WORKING'])
        }
        else if (c.value.text == 'https://nutrabay.com/'){
            ws2.addRow(['https://nutrabay.com/','WOOCOMMERCE'])
        }
        else if (c.value.text == 'https://shop.waaree.com/'){
            ws2.addRow(['https://shop.waaree.com/','BIGCOMMERCE'])
        }
        else if (c.value.text == 'https://www.cult.fit/store/gear'){
            ws2.addRow(['https://www.cult.fit/store/gear','OTHERS'])
        }
        else if (c.value.text == 'https://www.ritukumar.com/'){
            ws2.addRow(['https://www.ritukumar.com/','MAGENTO'])
        }
    });
    return wb.xlsx.writeFile('Zecko.xlsx');

}).catch(err => {
    console.log(`Error : ${err.message}`);
});