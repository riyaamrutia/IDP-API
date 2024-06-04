const axios = require('axios');
const xml2js = require('xml2js');
const XLSX = require('xlsx');
const moment = require('moment');

// Read Excel file
const workbook = XLSX.readFile('IDP.xlsx');
const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
const sheet = workbook.Sheets[sheetName];
const deliverableItemjson = XLSX.utils.sheet_to_json(sheet);

console.log("json data is: ", deliverableItemjson);

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


const formattedData = deliverableItemjson.map(row =>{
    const jsDate = excelDateToJSDate(row.plannedDate);
    return{
        ...row,
        plannedDate: moment(jsDate).format('DD-MMM-YYYY')
    };
});

console.log("Formatted data: ", formattedData);
// console.log("Excel deliverable Item json is :" ,deliverableItemjson);
// console.log("formatted deliverable Item json is :" ,formattedData);
const ans = formattedData[0];
//console.log("ans: ",ans);

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
//1
// Make the POST request using axios
axios.post("https://dmsak.asite.com/apilogin/", postData, {
    headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
})
    .then(response => {
        // Parse XML response to JSON
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
                        // Parse XML response to JSON
                        xml2js.parseString(response.data, (err, result) => {
                            if (err) {
                                console.error('Error parsing XML:', err);
                            } else {
                                let arr = [];
                                arr = result.asiteDataList.workspaceVO;

                                const ws_name = "IDP - Child - RA";
                                const ws_id = findWorkspaceId(ws_name, arr);
                                console.log("Project id: ", ws_id);

                                //3
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
                                    .then(response => {
                                        // reponse already in json
                                        if (err) {
                                            console.error('Error parsing XML:', err);
                                        } else {
                                            //FormList is object and form is array that represents a single form.
                                            const formMsgId = response.data.FormList.Form[0].FormMessages.FormMessage[0].FormMsgID;
                                            console.log("Form message id is: ", formMsgId);
                                            
                                            const deliverableItemString = JSON.stringify(ans);
                                            //console.log("jsonString: ", deliverableItemString);

                                            const postReqData = {
                                                projectId: ws_id,
                                                msgId: formMsgId,
                                                jsonData: deliverableItemString
                                            };
                                            console.log(postReqData);
                                            //4
                                            const cookies = response.headers['set-cookie'];
                                            console.log("Cookie: ", cookies);
                                            let cookieValue;
                                            // Loop through the array of cookies
                                            for (const cookie of cookies) {
                                                // Split the cookie string by ';' to get individual parts
                                                cookieValue = cookie.split(';');
                                                console.log("trimed cookie: ", cookieValue[0]);
                                                trimedCookie = cookieValue[0];
                                                // Loop through the parts to find the one that starts with 'JSESSIONID'

                                            }
                                            //2
                                            let trimmed = cookieValue[0];

                                            axios.post("https://adoddleak.asite.com/commonapi/htmlForm/create/deliverableItem", new URLSearchParams(postReqData),
                                                {
                                                    headers: {
                                                        'ASessionID': sessionId,
                                                        'Content-Type': 'application/x-www-form-urlencoded',
                                                        'Cookie': trimmed
                                                    },
                                                })
                                                .then(response => {
                                                    if (err) {

                                                        console.error('Error in Deliverable Creation:', err);
                                                    } else {
                                                        //console.log(response.data);
                                                        let siref = response.data.entity.siref;
                                                        console.log("Siref: ", siref);
                                                        console.log("Response message: ", response.data.message);
                                                        console.log("deliverable item created successfully");

                                                    //     let stageItemjson = JSON.stringify({
                                                    //         "projectCode": "EABIM001",
                                                    //         "organization": "ACM",
                                                    //         "volume": "00",
                                                    //         "location": "00",
                                                    //         "fileType": "AN",
                                                    //         "role": "A",
                                                    //         "fileNumber": "41",
                                                    //         "deliverableReference": "A0100",
                                                    //         "stage": "EA0",
                                                    //         "lOD": "LOD0",
                                                    //         "dTitle": "30 april",
                                                    //         "plannedDate": "30-Apr-2025",
                                                    //         "notifyTo": "drv.service@environment-agency.gov.uk",
                                                    //         "siref": siref
                                                    //     });

                                                    //     let stageData = {
                                                    //         projectId: ws_id,
                                                    //         msgId: formMsgId,
                                                    //         jsonData: stageItemjson
                                                    //     }

                                                    //     axios.post("https://adoddleak.asite.com/commonapi/htmlForm/create/stageItem", new URLSearchParams(stageData), {
                                                    //         headers: {
                                                    //             'ASessionID': sessionId,
                                                    //             'Content-Type': 'application/x-www-form-urlencoded',
                                                    //             'Cookie': trimmed
                                                    //         },
                                                    //     })
                                                    //         .then(response => {
                                                    //             if (err) {
                                                    //                 console.log('Error in stage item creation: ', err);
                                                    //             }
                                                    //             else {
                                                    //                 console.log(response.data);
                                                    //                 console.log("stage item created successfully");
                                                    //             }
                                                    //         })
                                                    //         .catch(error => {
                                                    //             console.error('Error:', error);
                                                    //         });
                                                     }
                                                })
                                                .catch(error => {
                                                    console.error('Error:', error);
                                                });
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

