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
  choices: ["getUsers", "getUser", "getDrive", "getDriveRoot"],
  demandOption: true,
}).argv;

const apiBase = process.env.GRAPH_ENDPOINT + "/v1.0";
async function main() {
  console.log(`You have selected: ${options.op}`, options);

  let result, error;

  // here we get an access token
  const authResponse = await auth.getToken(auth.tokenRequest);
  const { accessToken } = authResponse;

  try {
    switch (yargs.argv["op"]) {
      case "getUsers":
        const users = await fetch.callApi(`${apiBase}/users`, accessToken);
        result = users.value
          .filter((u) => !u.userPrincipalName.includes("#EXT#"))
          .map(({ displayName, userPrincipalName }) => ({
            displayName,
            userPrincipalName,
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

      case "getDriveRoot":
        const driveId =
          "b!tOvj5m3TxUuAQeyG1BWmJyPUbJ5L1slOqVB4aPdgGD2qfEMY5ue0R6EVDM71kgjB";

        result = await fetch.callApi(
          `${apiBase}/drives/${driveId}/root:/`,
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
