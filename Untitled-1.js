 .then(response => {
     // reponse already in json
     if (err) {
         console.error('Error parsing XML:', err);
     } else {
         //FormList is object and form is array that represents a single form.
         const formMsgId = response.data.FormList.Form[0].FormMessages.FormMessage[0].FormMsgID;
         console.log("Form message id is: ", formMsgId);

         const cookies = response.headers['set-cookie'];
         console.log("Cookie: ", cookies);
         let cookieValue;
         // Loop through the array of cookies
         for (const cookie of cookies) {
             // Split the cookie string by ';' to get individual parts
             cookieValue = cookie.split(';');
             //console.log("trimed cookie: ", cookieValue[0]);
             trimedCookie = cookieValue[0];
             // Loop through the parts to find the one that starts with 'JSESSIONID'

         }
         let trimmed = cookieValue[0];
         console.log("cookie: ", trimmed);

         const postRowData = async (rowData, rowIndex) => {

             console.log("row data is: " + rowIndex + "----" + rowData);

             const postReqData = {
                 projectId: ws_id,
                 msgId: formMsgId,
                 jsonData: JSON.stringify(rowData)//deliverableItemString
             };
             console.log("post data is: ", postReqData);

             try {
                 const response = axios.post("https://adoddleak.asite.com/commonapi/htmlForm/create/deliverableItem", new URLSearchParams(postReqData),
                     {
                         headers: {
                             'ASessionID': sessionId,
                             'Content-Type': 'application/x-www-form-urlencoded',
                             'Cookie': trimmed
                         },
                     });

                 let siref = response.data.entity.siref;
                 console.log("Siref: ", siref);
                 console.log("Response message: ", response.data.message);
                 console.log("deliverable item created successfully");

                 console.log(`Row ${rowIndex} posted successfully:`, response.data);
             } catch (error) {
                 console.error(`Error2 posting row ${rowIndex}:`, error.message);
             }
         };

         (async () => {
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
         })();
     }
 })
    .catch(error => {
        console.error('Error:', error);
    });