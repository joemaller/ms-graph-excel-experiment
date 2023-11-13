#!/usr/bin/env node

// read in env settings
require("dotenv").config();

const yargs = require("yargs");

const fetch = require("./fetch");
const auth = require("./auth");

const options = yargs.usage("Usage: --op <operation_name>").option("op", {
  alias: "operation",
  describe: "operation name",
  type: "string",
  choices: [
    "getUsers",
    "getUser",
    "getDrive",
    "getDriveChildren",
    "getSheetId",
    "getRows",
    "appendRows",
  ],
  demandOption: true,
}).argv;

/**
 * TODO: Move these into modules
 * @var date {Date} a date object, usually NOW
 */
const excelDate = (date) => {
  if (!date) {
    date = new Date();
  }
  date.setHours(0, 0, 0, 0); // set time to midnight
  zeroDay = new Date(1900, 0, 0);
  return (date - zeroDay) / (24 * 60 * 60 * 1000) + 1;
};

const excelTime = (date) => {
  if (!date) {
    date = new Date();
  }
  const midnight = new Date();
  midnight.setHours(0, 0, 0, 0);

  // NOTE: Excel rounds floating point values to 15 decimal places. JS sends a 16-digit floating point number
  return (date - midnight) / 86400000; // 24 hours * 60 minutes * 60 seconds * 1000 (ms)
};

const apiBase = process.env.GRAPH_ENDPOINT + "/v1.0";
async function main() {
  console.log(`You have selected: ${options.op}`, options);

  let result;

  let driveId =
    "b!tOvj5m3TxUuAQeyG1BWmJyPUbJ5L1slOqVB4aPdgGD2qfEMY5ue0R6EVDM71kgjB"; // iop@ideasonpurpose.com's OneDrive

  let sheetId = "01GEOOIVICYY5OWOVHEFHIGXW7WTF5JEFI"; // IOP/IOP021_ BNB/Brand New Brand! Cycle 5/a_Mangement/Contact Form Test Sheet.xlsx

  // here we get an access token
  const authResponse = await auth.getToken(auth.tokenRequest);
  const { accessToken } = authResponse;

  try {
    switch (yargs.argv["op"]) {
      case "getUsers":
        const users = await fetch.callApi(`${apiBase}/users`, accessToken);
        result = users.value
          .filter((u) => !u.userPrincipalName.includes("#EXT#"))
          .map(({ displayName, id, userPrincipalName }) => ({
            displayName,
            userPrincipalName,
            id,
          }));
        break;

      case "getUser":
        result = await fetch.callApi(
          `${apiBase}/users/${options._[0]}`,
          accessToken
        );
        break;

      case "getDrive":
        const driveUser = options._[0] ?? "iop@ideasonpurpose.com";
        result = await fetch.callApi(
          `${apiBase}/users/${driveUser}/drive`,
          accessToken
        );
        break;

      case "getDriveChildren":
        // const itemPath = 'Active Jobs';
        const itemPath =
          "Active Jobs/IOP/IOP021_ BNB/Brand New Brand! Cycle 5/a_Mangement";
        // https://ideasonpurpose-my.sharepoint.com/personal/iop_ideasonpurpose_com/Documents/Active%20Jobs
        // https://ideasonpurpose-my.sharepoint.com/personal/iop_ideasonpurpose_com/Documents/Active%20Jobs/IOP/IOP021_%20BNB/Brand%20New%20Brand!%20Cycle%205/a_Mangement/Contact%20Form%20Test%20Sheet.xlsx
        // NOTE: (big note) Check the webUrl to discover the effective root directory
        //       Requesting Users' OneDrive directories puts us in /Documents, eg.
        //           /personal/iop_ideasonpurpose_com/Documents
        //       This means any reqeusts for children should reference known folders
        //       No idea why the plain /root: doesn't' work
        result = await fetch.callApi(
          `${apiBase}/drives/${driveId}/root:/${itemPath}:/children`,
          // `${apiBase}/drives/${driveId}/root:/Active Jobs/IOP/IOP021_ BNB/Brand New Brand! Cycle 5/a_Mangement/Contact Form Test Sheet.xlsx`,
          accessToken
        );
        break;

      case "getSheetId":
        const filePath =
          "Active Jobs/IOP/IOP021_ BNB/Brand New Brand! Cycle 5/a_Mangement";
        const fileName = "Contact Form Test Sheet.xlsx";
        result = await fetch.callApi(
          `${apiBase}/drives/${driveId}/root:/${filePath}/${fileName}`,
          accessToken
        );
        console.log({ id: result.id });
        break;

      case "getRows":
        result = await fetch.callApi(
          `${apiBase}/drives/${driveId}/items/${sheetId}/workbook/worksheets/Sheet1/tables/Table1/rows`,
          accessToken
        );

        if (result.value.length) {
          result = result.value.map((r) => r.values);
        }

        break;

      case "appendRows":
        // TODO: Dates should be converted from the actual request timestamp
        const values = [
          [
            excelDate(),
            excelTime(),
            "name from API",
            "email@api.go",
            new Date(),
          ],
        ];
        result = await fetch.callApi(
          `${apiBase}/drives/${driveId}/items/${sheetId}/workbook/worksheets/Sheet1/tables/Table1/rows`,
          accessToken,
          "POST",
          JSON.stringify({ values })
        );
        break;

      default:
        result = "Select a Graph operation first";
        break;
    }

    console.log(result);
  } catch (error) {
    console.log(error);
  }
}

main();
