const express = require("express");

const { google } = require("googleapis");

const app = express();
app.set("view engine", "ejs");
app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.render("index");
});

app.post("/", async (req, res) => {
  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",

    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });

  // Create client instance for auth
  const client = await auth.getClient();

  // Instance of Google Sheets API
  const googleSheets = google.sheets({ version: "v4", auth: client });

  const spreadsheetId = "1RBICUYHFJoQD_saZ9OdFmnuR_EDC0a5uPiqBftzb_ds";

  // Get metadata about spreadsheet
  const metaData = await googleSheets.spreadsheets.get({
    auth,
    spreadsheetId,
  });

  // Get range of metadata
  const sheetRange = "engenharia_de_software";

  try {
    const situations = await calculateSituation(
      googleSheets,
      auth,
      spreadsheetId,
      sheetRange
    );

    processSituation(googleSheets, auth, spreadsheetId, situations, sheetRange);

    res.send("Médias calculadas com sucesso! Verfique a planilha");
  } catch (error) {
    res.status(500).send("Internal Server Error");
  }
});

app.listen(8000, (req, res) => console.log("running on 8000"));

async function calculateSituation(
  googleSheets,
  auth,
  spreadsheetId,
  sheetRange
) {
  // Array to store the situations
  const situations = [];

  const getRows = await googleSheets.spreadsheets.values.get({
    auth,
    spreadsheetId,
    range: `${sheetRange}!C4:F27`,
  });

  // Check if data.values is not undefined
  if (getRows.data.values) {
    // Convert the values to numbers
    const students = getRows.data.values.map((row) =>
      row.map((cell) => parseFloat(cell))
    );

    students.forEach((student, index) => {
      let absence = student[0];

      // Check if the student has reached or not the absence limit
      if (absence <= 15) {
        let scores = student.slice(1);

        // Calculate the average
        const sum = scores.reduce(function (acc, num) {
          return acc + num;
        }, 0);

        const average = Math.ceil(sum / scores.length);

        // Store the situations in the array
        situations.push(processAverage(average));
      } else {
        situations.push(["Reprovado por Falta", 0]);
      }
    });
  } else {
    // Handle the case where data.values is undefined
    console.log(`No data found for row ${i}`);
  }

  return situations;
}

async function processSituation(
  googleSheets,
  auth,
  spreadsheetId,
  situations,
  sheetRange
) {
  situations.forEach((student, index) => {
    const [situation, minimumPassingGrade] = student;

    // Write row(s) to spreadsheet
    googleSheets.spreadsheets.values.append({
      auth,
      spreadsheetId,
      range: `${sheetRange}!G${index + 4}:H${index + 4}`,
      valueInputOption: "USER_ENTERED",
      resource: {
        values: [[situation, minimumPassingGrade]],
      },
    });
  });
}

function processAverage(average) {
  let situation = "";
  let minimumPassingGrade = 0;

  if (average < 50) {
    situation = "Reprovado por Nota";
  }

  if (average >= 50 && average < 70) {
    minimumPassingGrade = 100 - average;
    situation = "Exame final";
  }

  if (average >= 70) {
    situation = "Aprovado";
  }

  return [situation, minimumPassingGrade];
}
