const express = require("express");
const {google} = require("googleapis");

const app = express();

const port = 3000;

app.get("/", async (req, res) => {
    // Google Auth authentication
    const auth = new google.auth.GoogleAuth({
        keyFile: "credentials.json", // This file contains the credentials you downloaded from your Google Sheets
        scopes: "https://www.googleapis.com/auth/spreadsheets",
    });

    const client = await auth.getClient();

    const googleSheets = google.sheets({version: "v4", auth: client});

    const spreadsheetId = 'YOUR_SPREADSHEET_ID'; // Link your Google spreadsheet by ID

    // Collect the data from the spreadsheet
    const getRows = await googleSheets.spreadsheets.values.get({
        auth,
        spreadsheetId,
        range: "engenharia_de_software!A2:F" // The name of your spreadsheet + the rows and columns to be read
    })

    // Extracting the data from the spreadsheet
    const values = getRows.data.values;

    // Checking the number of classes in the semester
    const classes = values[0][0].match(/\d+/g);

    // Initializing an array to store students' grades
    let situation = [];
    for (let i = 2; i < values.length; i++) {
        let sumGrades = 0;
        for (let j = 3; j <= 5; j++) {
            const grades = parseFloat(values[i][j]);
            if (!isNaN(grades)) {
                sumGrades += grades;
            }
        }

        // Calculate the final grade
        const finalGrade = Math.ceil((sumGrades / 10) / 3);

        // Check if the number of fouls
        if(values[i][2] > Math.ceil(classes * 0.25)){
            situation.push(["Reprovado por Falta", 0])
        }else{
            // Checking the student's situation
            if(finalGrade >= 7){
                situation.push(["Aprovado", 0]);
            }else if (finalGrade >= 5){
                const requiredGrade = 10 - finalGrade;
                situation.push(["Exame Final", requiredGrade]);
            }else{
                situation.push(["Reprovado por Nota", 0]);
            }
        }
    }
    
    console.log("Student's situation successfully extracted.");

    // Updating the data in the spreadsheet
    await googleSheets.spreadsheets.values.update({
        auth,
        spreadsheetId,
        range: "engenharia_de_software!G4:H", // The name of your spreadsheet + the rows and columns to be read
        valueInputOption: "USER_ENTERED",
        resource: {
            values: situation
        }
    });

    console.log("Spreadsheet successfully updated.");

    res.send(values);
})

app.listen(port, (req, res) => console.log("running on " + port));