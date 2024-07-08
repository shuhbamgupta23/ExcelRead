const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const axios = require("axios");
const qs = require("qs");
const https = require("https");
const CLIENT_ID = "9d408bfe-01d5-4cd2-ba1a-f1f9b4b53704";
const CLIENT_SECRET = "C1YhEkH7lSGJdBqx-CWc3w";
const API_BASE_URL = "https://spglobal-api-stg.unily.com";
const filePath = path.join(__dirname, "..", "Excel", "excel2.csv");
const logsFolderPath = path.join(__dirname, "..", "logs");

if (!fs.existsSync(logsFolderPath)) {
  fs.mkdirSync(logsFolderPath);
}

const workBook = xlsx.readFile(filePath);
const sheetName = workBook.SheetNames[0];
const workSheet = workBook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(workSheet);

let accessToken = null;

const getFormattedDate = () => {
  let fullDateTime = new Date();
  let options = {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  };

  let formattedDate = new Intl.DateTimeFormat("en-GB", options)
    .format(fullDateTime)
    .replace(",", "");
  let dateParts = formattedDate.split(" ")[0].split("/");

  let year = dateParts[2];
  let month = dateParts[1];
  let date = dateParts[0];

  return `${year}${month}${date}`;
};

const logResponse = (logData) => {
  const formattedDate = getFormattedDate();
  const logFilePath = path.join(logsFolderPath, `log_${formattedDate}.txt`);
  fs.appendFileSync(logFilePath, logData + "\n", "utf8");
};

const getAccessToken = async () => {
  const tokenAccessBody = qs.stringify({
    grant_type: "client_credentials",
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
  });

  const config = {
    method: "post",
    url: `${API_BASE_URL}/connect/token`,
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    data: tokenAccessBody,
    httpsAgent: new https.Agent({ rejectUnauthorized: false }),
  };

  try {
    const response = await axios(config);
    accessToken = response.data.access_token;
  } catch (error) {
    const logErrorData = `ErrorType: Fetching access token, Status Code: ${500}, Message: ${
      error.message
    }`;
    logResponse(logErrorData);
    console.error("Error fetching access token:", error.message);
  }
};

const retryWithTokenRefresh = async (
  config,
  userEmail,
  callback,
  retryCount = 0
) => {
  const maxRetries = 5;
  const retryDelay = 1000 * Math.pow(2, retryCount);

  try {
    const response = await axios(config);
    await callback(response);
  } catch (error) {
    if (error.response && error.response.status === 401) {
      const tokenExpiredLog = `Token expired for user ${userEmail}. Refreshing token and retrying...`;
      logResponse(tokenExpiredLog);
      console.log(tokenExpiredLog);
      await getAccessToken();
      config.headers.Authorization = `Bearer ${accessToken}`;
      try {
        const retryResponse = await axios(config);
        await callback(retryResponse);
      } catch (retryError) {
        const logErrorData = `ErrorType: Retry Error for user ${userEmail}, Status Code: ${500}, Message: ${
          retryError.message
        }`;
        logResponse(logErrorData);
        console.error(
          `Error retrying API call for user ${userEmail}:`,
          retryError.message
        );
      }
    } else if (
      error.response &&
      error.response.status === 503 &&
      retryCount < maxRetries
    ) {
      const retryLog = `Service unavailable for user ${userEmail}. Retrying in ${
        retryDelay / 1000
      } seconds... (Attempt ${retryCount + 1}/${maxRetries})`;
      logResponse(retryLog);
      console.log(retryLog);
      setTimeout(async () => {
        await retryWithTokenRefresh(
          config,
          userEmail,
          callback,
          retryCount + 1
        );
      }, retryDelay);
    } else {
      const logErrorData = `ErrorType: API Error for user ${userEmail}, Status Code: ${500}, Message: ${
        error.message
      }`;
      logResponse(logErrorData);
      console.error(
        `Error processing API call for user ${userEmail}:`,
        error.message
      );
    }
  }
};

const processEachEntry = async () => {
  for (let user of data) {
    await ensureValidToken();

    let currentUserRequestBody = {
      pageSize: 1,
      pagingToken: "",
      sort: {
        field: "createDate",
        direction: "Desc",
      },
      queryText: `email:${user.email}`,
    };

    let config = {
      method: "post",
      url: `${API_BASE_URL}/api/v1/users/query`,
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
      data: currentUserRequestBody,
      httpsAgent: new https.Agent({ rejectUnauthorized: false }),
    };

    await retryWithTokenRefresh(config, user.email, async (response) => {
      let properties = response.data.data[0]?.properties || [];

      let updatedProperties = properties.map((row) => {
        if (row.alias === "o365MigrationWaveDate") {
          row.value = user.year;
        }
        if (row.alias === "o365MigrationWave") {
          row.value = user.quarter;
        }
        if (row.alias === "o365MigrationCompleted") {
          row.value = user.persona;
        }
        return row;
      });

      let updatedResponse = response.data;
      updatedResponse.data.properties = updatedProperties;
      await updateEachEntry(updatedResponse.data);
    });
  }
};

const updateEachEntry = async (updatedValue) => {
  await ensureValidToken();

  let config = {
    method: "put",
    url: `${API_BASE_URL}/api/v1/users`,
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${accessToken}`,
    },
    data: updatedValue,
    httpsAgent: new https.Agent({ rejectUnauthorized: false }),
  };

  await retryWithTokenRefresh(
    config,
    updatedValue?.[0]?.email,
    async (response) => {
      console.log("Update successful:", response.data);
      const logData = `Email: ${updatedValue?.[0]?.email}, Status Code: ${
        response.data?.[0]?.statusCode
      }, Message: ${
        response.data?.[0]?.message === ""
          ? "Updation Successful"
          : response.data?.[0]?.message
      }`;
      logResponse(logData);
    }
  );
};

const ensureValidToken = async () => {
  if (!accessToken) {
    await getAccessToken();
  }
};

processEachEntry();
