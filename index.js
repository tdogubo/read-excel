const XLSX = require("xlsx");

// Function to convert the excel file to the json format for easy manipulation
function converter (item)  {
    let sheet2 = [];
    if (!item.length > 0) {
        return console.log("No data");
    }
    let keys = [];
    let values = [];
    
    for (const sheet of item) {
        keys.push(Object.keys(sheet).values().next().value.split("|"));
        values.push(Object.values(sheet).values().next().value.split("|"));
    }
    for (let i = 0; i < keys.length; i++) {
        sheet2.push(keys[i].reduce((acc, val, ind) => {
            acc[val] = values[i][ind];
            return acc;
        }, {}));
    }
    return sheet2;
};


// Function to convert the readFile and manulate based on condition
function editor() {
    const workbook = XLSX.readFile("file.xlsx");
    const file = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames]);
    
    let check = converter(file);

    const edited = check.map((prop) => {
        if (prop.Description.match("FEE CHG")) {
            prop.Flag = "FEE CHG";
        } else if (prop.Description.match("AEPS")) {
            prop.Flag = "AEPS";
        } else prop.Flag = "";

        [prop["Description "], prop[""]] = prop.Description.split("/");
        return prop;
    });

    const newBook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(edited);
    XLSX.utils.book_append_sheet(newBook, newSheet, "Edited Backend Test");
    XLSX.writeFile(newBook, "edited-file.xlsx");
}

editor();
