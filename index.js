const fs = require('fs');
const { utils, read } = require('xlsx');
const { google } = require('googleapis');

const dotenv = require('dotenv');
const { auth } = require('googleapis/build/src/apis/abusiveexperiencereport');

dotenv.config();

// Replace with your own Google Sheets template ID
const templateId = process.env.TEMPLATE_FILE_ID;
let rawCredentialdata = fs.readFileSync('credetials.json');
let credemtialdata = JSON.parse(rawCredentialdata);
const credentials = {
    client_email: credemtialdata.client_email,
    private_key: credemtialdata.private_key,
};

// Initialize the Google Sheets API client
const sheets = google.sheets('v4');

// Initialize the Google Drive API client
const drive = google.drive("v3");

async function getAuth() {
  const auth = new google.auth.JWT(credentials.client_email, null, credentials.private_key, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive.metadata']);
  return auth;
}

async function copyFileWithEditable(user_email, new_filename, source_file_id, auth) {
    // copy file
    const copiedFile = await drive.files.copy({
        fileId: source_file_id,
        requestBody: {
            name: new_filename,
            mimeType: 'application/vnd.google-apps.spreadsheet',
        },
        auth: auth
    }); 
    const newFileID = copiedFile.data.id;
    // add edit permission
    const res = await drive.permissions.create({
        resource: {
          type: "user",
          role: "writer",
          emailAddress: user_email,  // Please set the email address you want to give the permission.
        },
        fileId: newFileID,
        auth: auth
      });
    if(res.status == 200){
        const edit_url = `https://docs.google.com/spreadsheets/d/${newFileID}/edit#gid=0`;
        return {
            "fileID": newFileID,
            "URL": edit_url
        };
    } else {
        return res.status;
    }
}

async function writeDataInTemplateSheet( sheetID, data, auth) {
  try {
    const result = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetID,
      requestBody: {
        valueInputOption: "USER_ENTERED",
        data: data
      },
      auth: auth
    });
    return result;
  } catch (err) {
    console.log(err)
    throw err;
  }
}

function generateDataItem(sheetName, columnsID, startNum, data){
   const dataItem = {
    "range": `${sheetName}!${columnsID}${startNum}`,
    "majorDimension": "ROWS",
    "values": data
   }
   return dataItem;
}

function parseExceltoJson(filename, option, startNum) {
  const workbook = read(filename, { type: 'file' });
  const sheet_name_list = workbook.SheetNames[0];
  const rows = utils.sheet_to_json(workbook.Sheets[sheet_name_list], option);
  const ret_data = rows.slice(startNum, rows.length);
  return ret_data;
}

(async () => {
  
  const jwtClient = await getAuth();
  const new_file = await copyFileWithEditable("webhipe@gmail.com", "Test Formula-5", templateId, jwtClient);
  
  const ba_data = parseExceltoJson(
    'excel_data/ba.xls',
    {
      raw: true,
      dateNF: 'MM-DD-YYYY',
      header: 1,
      defval: '',
      range: 5,
    },
    1
  );
  
  const ba_date_data = Array(ba_data.length).fill(["01/05/2023"]);

  const attributes_data = parseExceltoJson(
    'excel_data/at.xls',
    {
      raw: true,
      dateNF: 'MM-DD-YYYY',
      header: 1,
      defval: '',
    },
    4
  )

  const data = [
    generateDataItem("RAW BA Data [MONTHLY]", "D", 2, ba_date_data),
    generateDataItem("RAW BA Data [MONTHLY]", "F", 2, ba_data),
    generateDataItem("Attributes", "A", 5, attributes_data)
  ]
  
  const result = await writeDataInTemplateSheet(new_file.fileID, data, jwtClient);
  console.log(new_file);
  console.log(result.status);

})();