
/**
 * Calls the endpoint with authorization bearer token.
 * @param {string} endpoint
 * @param {string} accessToken
 */
async function callApi(endpoint, accessToken) {
  const options = {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  };

  console.log({ "API Request": new Date().toString(), endpoint });

  try {
    const response = await fetch(endpoint, options);
    return await response.json();
  } catch (error) {
    console.log(error);
    return error;
  }
}

module.exports = {
  callApi: callApi,
};
