const axios = require('axios');
const xml2js = require('xml2js');
const XLSX = require('xlsx');
const moment = require('moment');

// format date
const excelDateToJSDate = (serial) => {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400; // seconds since Unix epoch                                        
    const date_info = new Date(utc_value * 1000);

    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    const total_seconds = Math.floor(86400 * fractional_day);

    const hours = Math.floor(total_seconds / 3600);
    const minutes = Math.floor((total_seconds % 3600) / 60);
    const seconds = total_seconds % 60;

    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
};

// Read Excel file
const workbook = XLSX.readFile('IDP.xlsx');
const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
const sheet = workbook.Sheets[sheetName];

// Get the range of rows and columns
const range = XLSX.utils.decode_range(sheet['!ref']);
console.log("range: ", range);

let sessionId;

// Define the data to be sent in the POST request
let postData = {
    emailId: 'ramrutiya@asite.com',
    password: 'Asite@123'
};

// find workspace id function
function findWorkspaceId(ws_name, arr) {
    for (const workspace of arr) {
        if (workspace.Workspace_Name == ws_name)
            return workspace.Workspace_Id;
    }
}

// Make the POST request using axios
axios.post("https://dmsak.asite.com/apilogin/", postData, {
    headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
})
    .then(response => {
        xml2js.parseString(response.data, (err, result) => {
            if (err) {
                console.error('Error parsing XML:', err);
            } else {
                sessionId = result.UserProfile.Sessionid;
                console.log('Session Id: ', sessionId);

                axios.get("https://dmsak.asite.com/api/workspace/workspacelist", {
                    headers: {
                        'ASessionID': sessionId,
                    }
                })
                    .then(response => {
                        xml2js.parseString(response.data, (err, result) => {
                            if (err) {
                                console.error('Error parsing XML:', err);
                            } else {
                                let arr = [];
                                arr = result.asiteDataList.workspaceVO;

                                const ws_name = "IDP - Child - RA";
                                const ws_id = findWorkspaceId(ws_name, arr);
                                console.log("Project id: ", ws_id);

                                axios.get("https://adoddleak.asite.com/commonapi/htmlForm/formSearch", {
                                    headers: {
                                        'ASessionID': sessionId,
                                    },
                                    params: {
                                        'ProjectName': ws_name,
                                        'FormCode': "IDP",
                                        'recordStart': "1",
                                        'recordLimit': "10",
                                        'includeMessages': "True"
                                    }
                                })
                                    .then(async response => {
                                        try {
                                            // Response already in JSON
                                            const formMsgId = response.data.FormList.Form[0].FormMessages.FormMessage[0].FormMsgID;
                                            console.log("Form message id is: ", formMsgId);

                                            const cookies = response.headers['set-cookie'];
                                            console.log("Cookie: ", cookies);
                                            let trimmed;
                                            // Loop through the array of cookies
                                            for (const cookie of cookies) {
                                                // Split the cookie string by ';' to get individual parts
                                                const cookieParts = cookie.split(';');
                                                // Extract the cookie value
                                                trimmed = cookieParts[0];
                                            }
                                            console.log("cookie: ", trimmed);

                                            const postRowData = async (rowData, rowIndex) => {
                                                console.log("row data is: " + rowIndex + "----" + rowData);

                                                const postReqData = {
                                                    projectId: ws_id,
                                                    msgId: formMsgId,
                                                    jsonData: JSON.stringify(rowData)
                                                };
                                                console.log("post data is: ", postReqData);

                                                try {
                                                    const postResponse = await axios.post(
                                                        "https://adoddleak.asite.com/commonapi/htmlForm/create/deliverableItem",
                                                        new URLSearchParams(postReqData),
                                                        {
                                                            headers: {
                                                                'ASessionID': sessionId,
                                                                'Content-Type': 'application/x-www-form-urlencoded',
                                                                'Cookie': trimmed
                                                            }
                                                        }
                                                    );

                                                    let siref = postResponse.data.entity.siref;
                                                    console.log("Siref: ", siref);
                                                    console.log("Response message: ", postResponse.data.message);
                                                    console.log("deliverable item created successfully");

                                                    console.log(`Row ${rowIndex} posted successfully:`, postResponse.data);
                                                } catch (error) {
                                                    console.error(`Error posting row ${rowIndex}:`, error.message);
                                                }
                                            };

                                            // Assuming the first row contains headers
                                            const headers = XLSX.utils.sheet_to_json(sheet, { header: 1, range: { s: { r: range.s.r, c: range.s.c }, e: { r: range.s.r, c: range.e.c } } })[0];

                                            for (let R = range.s.r + 1; R <= range.e.r; R++) {
                                                const row = XLSX.utils.sheet_to_json(sheet, { header: 1, range: { s: { r: R, c: range.s.c }, e: { r: R, c: range.e.c } } })[0];

                                                if (row && row[0] !== undefined) { // Ensure the row is not empty and has data
                                                    const rowData = {};

                                                    headers.forEach((header, index) => {
                                                        rowData[header] = row[index];
                                                    });

                                                    // Convert date if it exists in the rowData
                                                    if (rowData.plannedDate) {
                                                        const jsDate = excelDateToJSDate(rowData.plannedDate);
                                                        rowData.plannedDate = moment(jsDate).format('DD-MMM-YYYY');
                                                    }
                                                    // Post the data using Axios
                                                    await postRowData(rowData, R);
                                                }
                                            }
                                        } catch (error) {
                                            console.error('Error processing initial response:', error.message);
                                        }
                                    })
                                    .catch(error => {
                                        console.error('Error:', error);
                                    });
                            }
                        });
                    })
                    .catch(error => {
                        console.error('Error:', error);
                    });
            }
        });
    })
    .catch(error => {
        console.error('Error:', error);
    });

