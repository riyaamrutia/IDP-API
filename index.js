const xml2js = require('xml2js');
const querystring = require('querystring');
const qs = require('qs');
const axios = require('axios');

let deliverableItemjson = {
    "projectCode": "EABIM001",
    "organization": "ACM",
    "volume": "00",
    "location": "00",
    "fileType": "AN",
    "role": "A",
    "fileNumber": "121211",
    "deliverableReference": "A0100",
    "stage": "EA0",
    "lOD": "LOD0",
    "dTitle": "12 may",
    "plannedDate": "30-Apr-2025",
    "notifyTo": "ramrutiya@asite.com"
}

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
                                            const jsonString = JSON.stringify(deliverableItemjson);
                                            console.log("json data: ", jsonString);

                                            const postReqData = {
                                                projectId: ws_id,
                                                msgId: formMsgId,
                                                jsonData: JSON.stringify(deliverableItemjson)
                                            };
                                            console.log("post request data is : ", postReqData);
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

                                                        let stageItemjson = JSON.stringify({
                                                            "projectCode": "EABIM001",
                                                            "organization": "ACM",
                                                            "volume": "00",
                                                            "location": "00",
                                                            "fileType": "AN",
                                                            "role": "A",
                                                            "fileNumber": "43",
                                                            "deliverableReference": "A0100",
                                                            "stage": "EA0",
                                                            "lOD": "LOD0",
                                                            "dTitle": "12 may",
                                                            "plannedDate": "30-Apr-2025",
                                                            "notifyTo": "drv.service@environment-agency.gov.uk",
                                                            "siref": siref
                                                        });

                                                        let stageData = {
                                                            projectId: ws_id,
                                                            msgId: formMsgId,
                                                            jsonData: stageItemjson
                                                        }

                                                        axios.post("https://adoddleak.asite.com/commonapi/htmlForm/create/stageItem", new URLSearchParams(stageData), {
                                                            headers: {
                                                                'ASessionID': sessionId,
                                                                'Content-Type': 'application/x-www-form-urlencoded',
                                                                'Cookie': trimmed
                                                            },
                                                        })
                                                            .then(response => {
                                                                if (err) {
                                                                    console.log('Error in stage item creation: ', err);
                                                                }
                                                                else {
                                                                    console.log(response.data);
                                                                    console.log("stage item created successfully");
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

