/*****************************************************************************************
 * function createKuttLink(kuttApiUrl, kuttApiKey, originalUrl, desiredShortUrl)
 * 
 * A function to create shortlinks from URLs using the Kutt API
 * 
 * This function requires you to pass in the following parameters:
 * @param {string} kuttApiUrl The full URL for the Kutt v2 API links endpoint (e.g., "https://kutt.yourdomain.com/api/v2/links").
 * @param {string} kuttApiKey The API key for authentication with Kutt.
 * @param {string} originalUrl The long URL to be shortened (maps to 'target').
 * @param {string} [desiredShortUrl] Optional. The custom short URL string to use (maps to 'customurl').
 * @param {string} [description] Optional. A description for the link (maps to 'description').
 * @param {string} [expiration] Optional. How long the link should last (e.g., "2 days", "1 hour") (maps to 'expire_in').
 * @param {string} [password] Optional. Password to protect the link (maps to 'password').
 * @param {boolean} [reuse] Optional. Whether to reuse an existing short URL (maps to 'reuse').
 * @param {string} [customDomain] Optional. The custom domain to use for the link (maps to 'domain').
 * 
 * The function returns the created Kutt short URL 'link' field)  on success or an error message string.
 * * @returns {string} url or error
******************************************************************************************/
 
function createKuttLink(kuttApiUrl, kuttApiKey, originalUrl, desiredShortUrl, description, expiration, password, reuse, customDomain) {
  // Input validation for required fields
  if (!kuttApiUrl || kuttApiUrl.trim() === "") {
    return "Error: Kutt API URL is required.";
  }
  if (!kuttApiKey || kuttApiKey.trim() === "") {
    return "Error: Kutt API Key is required.";
  }
  if (!originalUrl || originalUrl.trim() === "") {
    return "Error: Original URL is required.";
  }

  // Construct the payload based on the provided API schema
  const payload = {
    target: originalUrl.trim()
  };

  // Add optional fields if they are provided and not empty
  if (desiredShortUrl && desiredShortUrl.trim() !== "") {
    payload.customurl = desiredShortUrl.trim();
  }
  if (description && description.trim() !== "") {
    payload.description = description.trim();
  }
  if (expiration && expiration.trim() !== "") {
    payload.expire_in = expiration.trim();
  }
  if (password && password.trim() !== "") {
    payload.password = password.trim();
  }
  // 'reuse' is a boolean, convert string "TRUE" to true boolean
  if (typeof reuse === 'boolean') {
    payload.reuse = reuse;
  } else if (typeof reuse === 'string' && reuse.trim().toLowerCase() === 'true') {
    payload.reuse = true;
  } else {
    payload.reuse = false; // Default to false if not explicitly true
  }
  if (customDomain && customDomain.trim() !== "") {
    payload.domain = customDomain.trim();
  }

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "X-API-Key": kuttApiKey.trim()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Crucial to catch API errors without stopping the script
  };

  try {
    const response = UrlFetchApp.fetch(kuttApiUrl.trim(), options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 201) { // 201 Created is the success status for this API
      const responseJson = JSON.parse(responseText);
      // As per the provided schema, the short URL is in the 'link' field
      if (responseJson && responseJson.link) {
        Logger.log(`Successfully created shortlink for ${originalUrl}: ${responseJson.link}`);
        return responseJson.link; // Return the short URL
      } else {
        // This case indicates a 201 success but unexpected JSON structure
        Logger.log(`Kutt API Success (201) but unexpected response format for ${originalUrl}: ${responseText}`);
        return `API Success (201) but unexpected format: ${responseText}`;
      }
    } else {
      // Handle non-201 HTTP error codes or Kutt internal errors
      let errorMessage = `HTTP Error ${responseCode}`;
      try {
        const errorJson = JSON.parse(responseText);
        // Kutt errors often come with a 'message' field directly
        if (errorJson.message) {
          errorMessage += `: ${errorJson.message}`;
        } else {
          // Fallback if 'message' is not present, use raw response text
          errorMessage += `: ${responseText}`;
        }
      } catch (e) {
        // If response is not valid JSON, use raw text
        errorMessage += `: ${responseText}`;
      }
      Logger.log(`HTTP Error for ${originalUrl}: ${errorMessage}`);
      return `HTTP Error: ${errorMessage}`;
    }
  } catch (e) {
    // General script execution error (network, parsing, etc.)
    Logger.log(`Script Error for ${originalUrl}: ${e.message}`);
    return `Script Error: ${e.message}`;
  }
}


/**
 * Example function to demonstrate how to use createKuttLink for bulk processing.
 * This function reads from the active sheet, processes links one by one,
 * and writes the result back.
 */
function bulkCreateKuttLinksFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Assuming headers are in row 1, data starts from row 2
  const headers = values[0];
  const originalUrlCol = headers.indexOf("Original URL");
  const desiredShortUrlCol = headers.indexOf("Desired Short URL");
  const descriptionCol = headers.indexOf("Description");
  const expirationCol = headers.indexOf("Expiration");
  const passwordCol = headers.indexOf("Password");
  const reuseCol = headers.indexOf("Reuse Existing");
  const customDomainCol = headers.indexOf("Custom Domain");
  const kuttApiUrlCol = headers.indexOf("Kutt API URL");
  const kuttApiKeyCol = headers.indexOf("Kutt API Key");
  const resultCol = headers.indexOf("Kutt Short URL / Status");

  // Basic check for required columns
  if (originalUrlCol === -1 || kuttApiUrlCol === -1 || kuttApiKeyCol === -1 || resultCol === -1) {
    SpreadsheetApp.getUi().alert("Error: One or more required columns ('Original URL', 'Kutt API URL', 'Kutt API Key', 'Kutt Short URL / Status') not found in the active sheet. Please ensure column headers match exactly.");
    return;
  }

  // Iterate through rows, starting from the second row (index 1)
  for (let i = 1; i < values.length; i++) {
    const originalUrl = values[i][originalUrlCol];
    const desiredShortUrl = (desiredShortUrlCol !== -1) ? values[i][desiredShortUrlCol] : "";
    const description = (descriptionCol !== -1) ? values[i][descriptionCol] : "";
    const expiration = (expirationCol !== -1) ? values[i][expirationCol] : "";
    const password = (passwordCol !== -1) ? values[i][passwordCol] : "";
    const reuse = (reuseCol !== -1) ? values[i][reuseCol] : false; // Default to false
    const customDomain = (customDomainCol !== -1) ? values[i][customDomainCol] : "";

    const kuttApiUrl = values[i][kuttApiUrlCol];
    const kuttApiKey = values[i][kuttApiKeyCol];
    const currentResult = values[i][resultCol]; // Check if already processed

    // Skip if original URL is empty or result column is already populated with a valid URL
    if (!originalUrl || originalUrl.trim() === "" || (currentResult && currentResult.trim().startsWith("http"))) {
      continue;
    }

    // Call the core function for each row, passing all parameters
    const result = createKuttLink(
      kuttApiUrl,
      kuttApiKey,
      originalUrl,
      desiredShortUrl,
      description,
      expiration,
      password,
      reuse,
      customDomain
    );

    // Write the result back to the sheet
    sheet.getRange(i + 1, resultCol + 1).setValue(result); // +1 because getRange is 1-indexed

    // Add a small delay to avoid hitting rate limits if you have many links
    Utilities.sleep(500); // 0.5 second delay
  }
  SpreadsheetApp.getUi().alert("Kutt link creation process finished.");
}


/*****************************************************************************************
 * function checkUrlStatus(url)
 * 
 * A function validate URLs
 * 
 * This function requires you to pass in the following parameters:
 *  * @param {string} url
 * 
 * The function returns the status of the link.
 * * @returns {string} status
******************************************************************************************/

function checkUrlStatus(url) {
  if (!url || typeof url !== 'string' || url.trim() === '') {
    return "INVALID URL INPUT";
  }
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true // Important to follow redirects
    });
    var responseCode = response.getResponseCode();
    if (responseCode >= 200 && responseCode < 300) {
      return "LIVE";
    } else if (responseCode >= 300 && responseCode < 400) {
      return "REDIRECT (" + responseCode + ")";
    } else if (responseCode >= 400 && responseCode < 500) {
      return "CLIENT ERROR (" + responseCode + ")";
    } else if (responseCode >= 500 && responseCode < 600) {
      return "SERVER ERROR (" + responseCode + ")";
    } else {
      return "UNKNOWN STATUS (" + responseCode + ")";
    }
  } catch (e) {
    return "ERROR: " + e.message;
  }
}


/*****************************************************************************************
 * function urlencode(astr)
 * 
 * A function to safely encode strings for use within URL components
 * It prevents issues with special characters.
 * 
 * This function requires you to pass in the following parameters:
 *  * @param {string} url
 * 
 * The function the encoded string
 * * @returns {string} encodedURL
******************************************************************************************/

function urlencode(astr) {
  return encodeURIComponent(astr)
}



/*****************************************************************************************
 * function bitlyShortenUrlOriginal(longLink, accessToken)
 * 
 * A function to safely encode strings for use within URL components
 * It prevents issues with special characters.
 * 
 * This function requires you to pass in the following parameters:
 *  * @param {string} longLink The URL to shorten
 *  * @param {string} bitlyAPIKey The API key for authentication with Bitly.
 * 
 * You can change the domain to use if you have a custom domain registered with Bitly.
 *   var form = { 
    "domain": "bit.ly", (change to your custom domain here)
    "long_url": longLink
  };
  In practice, you can duplicate the function to call whichever custom domain you need.
 * 
 * The function returns the shortened URL
 * * @returns {string} url or error
******************************************************************************************/

function bitlyShortenUrlOriginal(longLink, accessToken) {
  var form = { 
    "domain": "bit.ly",
    "long_url": longLink
  };
  var url = "https://api-ssl.bitly.com/v4/shorten";

  var options = {
      "headers": { Authorization: `Bearer ${accessToken}` },
      "method": "POST",
      "contentType" : "application/json",
      "payload": JSON.stringify(form)
  };
  var response = UrlFetchApp.fetch(url, options);
  var responseParsed = JSON.parse(response.getContentText());
  return(responseParsed["link"]);
}

function bitlyShortenUrl(longLink, accessToken) {
  var form = { 
    "domain": "in.isesaki.in",
    "long_url": longLink
  };
  var url = "https://api-ssl.bitly.com/v4/shorten";

  var options = {
      "headers": { Authorization: `Bearer ${accessToken}` },
      "method": "POST",
      "contentType" : "application/json",
      "payload": JSON.stringify(form),
      'muteHttpExceptions': false
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var responseParsed = JSON.parse(response.getContentText());
  return(responseParsed["link"]);
}



/*****************************************************************************************
 * function ExpandURL(url) 
 * 
 * A function to expand a shortened URL by fetching the initial response
 * and reading the 'Location' header for the redirect URL.
 * 
 * NOTE: This only gets the FIRST redirect hop.
 * NOTE 2:This was made to pull long links out of Google's deprecated shortener
 * (It will likely NOT work for goo.gl links after Aug 25, 2025.)
 * 
 * This function requires you to pass in the following parameters:
 * @param {string} url The shortened URL.
 * @return {string} The URL from the Location header, or an error message.
 * @customfunction
 * 
 * The function returns the URL from the Location header, or an error message.
 * @return {string} url or error
******************************************************************************************/

function ExpandURL(url) {
  if (!url) {
    return "Please provide a URL";
  }

  try {
    // Fetch the URL without following redirects
    var response = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });

    // Get all headers from the response
    var headers = response.getHeaders();

    // Check if a 'Location' header exists (typical for redirects)
    if (headers['Location']) {
      // Decode URL-encoded characters and return the location
      // Need to handle cases where Location might be an array if multiple headers exist
      var locationUrl = Array.isArray(headers['Location']) ? headers['Location'][0] : headers['Location'];
      return decodeURIComponent(locationUrl);
    } else {
      // No Location header usually means no redirect, or an error occurred
      var responseCode = response.getResponseCode();
      if (responseCode >= 200 && responseCode < 300) {
         return "URL did not redirect (Status: " + responseCode + ")";
      } else {
         return "Error fetching URL or no redirect header (Status: " + responseCode + ")";
      }
    }

  } catch (e) {
    // Catch any other unexpected errors
    return "Error: " + e.toString();
  }
}