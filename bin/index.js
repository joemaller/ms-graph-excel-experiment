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
  choices: ["getUsers", "getUser", "getDrive", "getDriveChildren", "getSheet"],
  demandOption: true,
}).argv;

const apiBase = process.env.GRAPH_ENDPOINT + "/v1.0";
async function main() {
  console.log(`You have selected: ${options.op}`, options);

  let result, driveId;

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
        driveId =
          "b!tOvj5m3TxUuAQeyG1BWmJyPUbJ5L1slOqVB4aPdgGD2qfEMY5ue0R6EVDM71kgjB";

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

      case "getSheet":
        driveId =
          "b!tOvj5m3TxUuAQeyG1BWmJyPUbJ5L1slOqVB4aPdgGD2qfEMY5ue0R6EVDM71kgjB";

        const filePath =
          "Active Jobs/IOP/IOP021_ BNB/Brand New Brand! Cycle 5/a_Mangement";
        const fileName = "Contact Form Test Sheet.xlsx";
        result = await fetch.callApi(
          `${apiBase}/drives/${driveId}/root:/${filePath}/${fileName}`,
          accessToken
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
