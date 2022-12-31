let today = new Date("2022-12-7");
const month = today.toLocaleString("default", { month: "short" }).toUpperCase();

async function fetchProjects(name) {
  const isWeekend = today.getDay() === 0 || today.getDay() === 6;
  const white = { red: 1, green: 1, blue: 1 };

  const naProject = { key: "N/A", value: { value: "N/A", color: white } };
  const weekendProject = {
    key: "Weekend",
    value: { value: "🥳", color: {} },
  };

  if (isWeekend) {
    return { am: weekendProject, pm: weekendProject };
  }

  const { am, pm } = await getScheduledEntry(name);
  return { am: am ?? naProject, pm: pm ?? naProject };
}

/**
 * Disables the fetch button
 * @param {boolean} disable - whether to disable the button, true by default
 */
function disableFetchButton(disable = true) {
  document.getElementById("fetch-btn").disabled = disable;
}

async function renderSchedule() {
  const name = document.getElementById("name-input").value;
  if (!name) {
    renderError("Please enter your name");
    return;
  }

  disableFetchButton();

  renderError("");
  renderInfo("Fetching schedule... (0/3)");

  let amProject, pmProject;
  try {
    ({ am: amProject, pm: pmProject } = await fetchProjects(name));
  } catch (e) {
    console.error(e);
    renderError(e.message ?? "Something went wrong");
    return;
  } finally {
    disableFetchButton(false);
  }

  renderInfo("Fetching schedule... (3/3)");

  const amTextDisplay = document.getElementById("am-text");
  const amColorDisplay = document.getElementById("am-color");

  const pmTextDisplay = document.getElementById("pm-text");
  const pmColorDisplay = document.getElementById("pm-color");

  renderProject(amProject, amTextDisplay, amColorDisplay);
  renderProject(pmProject, pmTextDisplay, pmColorDisplay);

  renderInfo(`Schedule for ${name} on ${today.toDateString()}`);
}

/**
 * Renders the project on the page
 * @param {{key: string, value: {value: string, color: any}}} project - project object
 * @param {HTMLDivElement} textElem - element to render the project name
 * @param {HTMLDivElement} colorElem - element to render the project color
 */
function renderProject(project, textElem, colorElem) {
  const { key } = project;
  textElem.innerText = key;

  const { value, color } = project.value;
  const cssColor = {
    red: (color.red ?? 0) * 255,
    green: (color.green ?? 0) * 255,
    blue: (color.blue ?? 0) * 255,
  };
  colorElem.innerText = value;
  colorElem.style.backgroundColor = `rgb(${cssColor.red}, ${cssColor.green}, ${cssColor.blue})`;
}

/**
 * Fetch the project that the person is scheduled for, for both AM and PM
 * @param {string} name - name of the person
 * @returns {Promise<{am: {key: string, value: {value: string, color: any}}, pm: {key: string, value: {value: string, color: any}}}>} - object containing the AM and PM projects
 */
async function getScheduledEntry(name) {
  if (!access_token) {
    throw new Error("No access token");
  }

  const [rowNumber, columnNumber, keys] = await Promise.all([
    getRowNumber(name),
    getColumnNumber(),
    getKeys(),
  ]);

  renderInfo("Fetching schedule... (1/3)");

  if (columnNumber === -1 || rowNumber === -1) {
    return { am: undefined, pm: undefined };
  }

  const [cellAM, cellPM] = await getCell(rowNumber, columnNumber);
  const isHoliday = (cell) => Object.keys(cell.color).length === 0;

  renderInfo("Fetching schedule... (2/3)");

  let scheduleAM = keys.find(
    (k) =>
      k.value.value === cellAM.value && _.isEqual(k.value.color, cellAM.color)
  );
  let schedulePM = keys.find(
    (k) =>
      k.value.value === cellPM.value && _.isEqual(k.value.color, cellPM.color)
  );

  if (isHoliday(cellAM)) {
    scheduleAM = {
      key: "Holiday",
      value: { value: cellAM.value, color: {} },
    };
  }
  if (isHoliday(cellPM)) {
    schedulePM = {
      key: "Holiday",
      value: { value: cellPM.value, color: {} },
    };
  }

  return {
    am: scheduleAM,
    pm: schedulePM,
  };
}

/**
 * Fetches the keys from the key sheet
 * @param {string} name - name of the person
 * @returns {Promise<{key: string; value: {value: string, color: any}}[]>} - array of keys
 */
async function getKeys() {
  const ranges = [`KEY!A2:B31`, "KEY!A34:B48", "KEY!D2:E9"];

  const { data } = await googleSheetGET({
    query: [
      ["ranges", ranges[0]],
      ["ranges", ranges[1]],
      ["ranges", ranges[2]],
      [
        "fields",
        "sheets(data(rowData(values(effectiveValue,effectiveFormat))))",
      ],
    ],
  });
  const keysData = data.sheets[0].data;

  const keys = keysData
    .map((d) =>
      d.rowData.map((rd) => ({
        key: rd.values[0].effectiveValue.stringValue,
        value: {
          value: rd.values[1].effectiveValue?.stringValue ?? "",
          color: rd.values[1].effectiveFormat?.backgroundColor,
        },
      }))
    )
    .flat();

  return keys;
}

/**
 * @param {number} row
 * @param {number} col
 * @returns {Promise<{value: string, color: {red: number, green: number, blue: number}}[]>}
 */
async function getCell(row, col) {
  const colLetterAM = columnNumberToIndex(col);
  const colLetterPM = columnNumberToIndex(col + 1);

  const range = `${month}!${colLetterAM}${row}:${colLetterPM}${row}`;
  const { data } = await googleSheetGET({
    query: {
      ranges: range,
      fields: "sheets.data.rowData.values(effectiveFormat,effectiveValue)",
    },
  });
  const cells = data.sheets[0].data[0].rowData[0].values;

  return cells.map((c) => ({
    value: c.effectiveValue?.stringValue ?? "",
    color: c.effectiveFormat.backgroundColor,
  }));
}

/**
 * Fetches the row index corresponding to the person's name
 * @param {string} name Name of the person
 * @returns {Promise<number>} The row number corresponding to the person's name or -1 if not found
 */
async function getRowNumber(name) {
  const sheet = month;
  const range = `${sheet}!B:B`;

  const { data } = await googleSheetGET({
    endpoint: `values/${range}`,
    query: { majorDimension: "COLUMNS" },
  });
  const names = data.values;

  const nameIndex = names[0].findIndex(
    (n) => n?.toLowerCase() === name.toLowerCase()
  );

  return nameIndex === -1 ? null : nameIndex + 1;
}

/**
 * Fetches the column index corresponding to today's date
 * @returns {Promise<number>} The column number corresponding to today's date or -1 if not found
 */
async function getColumnNumber() {
  const sheet = month;
  const range = `${sheet}!2:2`;

  const { data } = await googleSheetGET({
    query: {
      ranges: range,
      fields: "sheets.data.rowData.values.effectiveValue",
    },
  });
  const dates = data.sheets[0].data[0].rowData[0].values;

  const index = dates.findIndex(
    (d) => d?.effectiveValue?.numberValue === today.getDate()
  );

  return index === -1 ? null : index + 1;
}

/**
 * Converts a column number to an alphabetical index
 * @param {number} columnNumber
 * @returns {string} The aphabetical column index
 */
function columnNumberToIndex(columnNumber) {
  if (columnNumber <= 0) {
    return "";
  }
  const remainder = (columnNumber - 1) % 26;
  return (
    columnNumberToIndex(Math.floor((columnNumber - 1) / 26)) +
    String.fromCharCode(65 + remainder)
  );
}

function renderInfo(message) {
  document.getElementById("info").innerText = message;
}

function renderError(message) {
  document.getElementById("error").innerText = message;
}

async function googleSheetGET({ endpoint, query }) {
  const spreadsheetId = "1-l37wl_YlE6AsL_ao4nHxs1ooIxpuRUwwIjfzWl82m4";
  const baseSheetsUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;

  const queryParams = new URLSearchParams(query);
  const url = `${baseSheetsUrl}/${endpoint ?? ""}?${queryParams.toString()}`;

  return axios.get(url, {
    headers: { Authorization: `Bearer ${access_token}` },
  });
}
