/**
 * Calls the endpoint with authorization bearer token.
 * @param {string} endpoint
 * @param {string} accessToken
 */
export async function callApi(endpoint, accessToken, method = "GET", body) {
  const options = {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    body
  };

  if (method === "POST") {
    options.headers["Content-Type"] = "application/json";
  }

  console.log({ "API Request": new Date().toString(), endpoint });

  try {
    const response = await fetch(endpoint, options);
    return await response.json();
  } catch (error) {
    console.log(error);
    return error;
  }
}

// module.exports = {
//   callApi: callApi,
// };
