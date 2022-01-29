
// shorthand for document.querySelector
const qry = (str) => {
    return document.querySelector(str)
}
const qryAll = (str) => {
    return document.querySelectorAll(str)
}


/**************************************/
/*                                    */
/*       Spreadsheet functions        */
/*                                    */
/**************************************/

// This loads the spreadsheet into the sheet js function and allows us to interact with it. 
const loadSpreadSheet = (spreadsheet) => {
    workbook = XLSX.read(spreadsheet, {cellDates:true});
    return workbook
}


// This accepts a workbook, opens a specific sheet, and converts each row into a json object.

const getSingleSheetJSON = (workbook, sheetNumber = 0, headerRow = 0) => {
    let sheet = workbook.Sheets[workbook.SheetNames[sheetNumber]]
    
    // range sets the header row
    // deval ensures that empty columns are also returned. 
    let loadedSheet = XLSX.utils.sheet_to_json(sheet, { range:headerRow, defval: ""})
    
    return loadedSheet
}


/*
Accepts two arguments.

1. newObjectMap - an object specifying which keys need to be mapped and renamed. 
{ 
    key1: newKey1,
    key2: newKey2 
}

2. The worksheet with all the data.

Returns a new list of objects with the new keys. 

[
    { newkey1: key1 value }, 
    { newkey1: key2 value }
]

*/

const transformWorksheet = (newObjectMap, worksheet) => {
    const newObjects = []
    worksheet.forEach(function(row, index){
        let keys = Object.keys(newObjectMap)
        const obj = {}
        for (let key of keys){
            obj[newObjectMap[key]] = row[key]
        }
        newObjects.push(obj)
    })
    return newObjects
}


/***************************************/
/*                                     */
/*          Global variables           */
/*                                     */
/***************************************/

// These store the loaded spreadsheet so that the other functions can manipulate them

let workbook, worksheet



// gets spreadsheet from local os
const getFile = async () => {
    const fileUpload = qry('#fileUpload').files[0]
    const data = await fileUpload.arrayBuffer();
    workbook = await loadSpreadSheet(data)
    loadStepTwo()
}

/*
1. If the loaded spreadsheet has multiple pages, the user is given an option to specify which sheet to load. 
2. The user is asked to specify which row is the header row. The default is the first row. 
*/
const loadStepTwo = () => {
    const numberOfSheets = workbook.SheetNames.length
    let res = '<h2>Step 2</h2><div class="inputGroup">'
    if (numberOfSheets > 1){
        let pickSheet = '<label for="pickSheet">Which sheet to load?</label><select id="pickSheet"><option selected value = 1>1</option>'
        for (var i = 2; i <= numberOfSheets; i++){
            pickSheet += `<option value=${i}>${i}</option>`
        }
        pickSheet += '</select>'
        res += pickSheet
    }
    let chooseHeader = '<label for="chooseHeader">Which row is the header?</label><input type="text" id="chooseHeader" placeholder = "default = 1" />'
    res += chooseHeader
    res += '<div class="width-100"><button id="stepTwoSubmit" type="button" onClick = "submitStepTwo()"role="button">submit</button></div></div>' 
    showResult(res, '#stepTwo')
    
}

// Submits the chosen sheet and header row and renders a preview of the object.
const submitStepTwo = () => {

    let sheetNumber = qry('#pickSheet') ? qry('#pickSheet').value - 1 : 0
    let headerRow = qry('#chooseHeader').value ? qry('#chooseHeader').value - 1 : 0

    // let loadedSheet = workbook.Sheets[workbook.SheetNames[sheet]]
    // worksheet = XLSX.utils.sheet_to_json(loadedSheet, { range:range-1, defval: ""})

    worksheet = getSingleSheetJSON(workbook, sheetNumber, headerRow )

    loadInitialPreview()
    loadStepThree()
}


// renders the row preview
const loadInitialPreview = () => {
    let row = visualizeRow(worksheet[0], 0, worksheet.length-1, 'old', '#initialPreview')
    let preview = '<h2>Preview</h2>'
    showResult(preview, '#stepTwoTitle')
    showResult(row, '#initialPreview')
}

// Renders a series of inputs for every key in the current object. 
// Here the user specifies which rows they want to import and gives them a name. 
const loadStepThree = () => {
    let row = worksheet[0]
    let keys = Object.keys(row)
    let map = '<table>'

    keys.forEach(function(key, index){
        let div = `
            <tr class="newObject">
                <td style="padding-right: 8px"><strong>${key}</strong></td>
                <td><input class="newAttribute" type="text" data-target="${key}" placeholder="Set key for this attribute. Leave blank to ignore"/></td>
            </tr>
        `
        map += div
    })
    showResult('<h2>Step 3 - Create your object structure </h2>', '#stepThree')
    map += '</table><div><button onclick="createNewObject()">submit</button>'
    showResult(map, '#newObject')

}

// placeholder for the newly created object. 
let newObjects = []

// processes the user input and then calls the transformWorsheet function to create the new object representation. 
const createNewObject = () => {
    const newObjectMap = {}
    newObjects = []
    
    newAttributes = qryAll('.newAttribute')
    newAttributes.forEach(function(item, index){
        if (item.value){
            newObjectMap[item.dataset.target] = item.value
        }
    })
    console.log(newObjectMap)

    newObjects = transformWorksheet(newObjectMap, worksheet)
    loadStepFour()
}

// renders a preview of a single row from the newly created object. 
const loadStepFour = () => {
    let row = visualizeRow(newObjects[0], 0, newObjects.length-1, 'new', '#newObjectPreview')
    let preview = '<h2>Step 4 - Preview new object</h2>'
    showResult(preview, '#stepFour')
    showResult(row, '#newObjectPreview')
}

// Generates a preview of the json
const generateJSON = () => {
    let res = `
        <h2>Step 5 - your data</h2>
        <button onclick="copyToClipboard()" type="button">Copy json to clipboard</button>
        <div id="jsonData">
            <pre>${JSON.stringify(newObjects,null, 2)}</pre>
        </div>
        `
    showResult(res, '#stepFive')
}

// copies json to clipboard
const copyToClipboard = () => {
    let data = JSON.stringify(newObjects,null, 2)
    navigator.clipboard.writeText(data)
}

// This navigates the rows in the spreadsheet / new object for previewing the data. 
const iterateRows = (row, sheet, destination) => {
    let ws = sheet == 'old' ?  worksheet : newObjects
    showResult(visualizeRow(ws[row], row, ws.length - 1, sheet, destination), destination)
}

// turns a row in a spreadsheet or an object into a table element for previewing. 
const visualizeRow = (row, index, max, sheet, destination) =>{
    res = ''
    for (let [key, value] of Object.entries(row)){
        res += `<tr><td><strong>${key}:</strong></td><td>${value}</td></tr>`
    }
    res = '<table>' + res + '</table>'
    let prev = index > 0 
        ? `<button role="button" type="button" onclick="iterateRows(${index-1},'${sheet}','${destination}')">Previous</button>`
        : '<button disabled>Previous</button>'
    let next = index < max 
        ? `<button role="button" type="button" onclick="iterateRows(${index+1}, '${sheet}', '${destination}')">Next</button>`
        : '<button disabled>Disabled</button>'

    let generate = sheet == 'new'
        ? '<button style="margin-left:16px" type = "button" onclick="generateJSON()">Generate json </button>'
        : ''
    
    let nav = '<div class="nav">'+ prev + next + generate + '</div>'
    res += nav
    return res
}

// renders text to a destination. 
const showResult = (val, destination) => {
    qry(destination).innerHTML = `<div>${val}</div>`
}

