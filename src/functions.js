const xlsx = require("xlsx");
const fs = require("fs/promises");
const path = require("path");
const config = require("../config.json");

function isValid(NGWordsArray, target) {
  for (const NGWord of NGWordsArray) {
    if (target.indexOf(NGWord) !== -1) return false;
  }
  return true;
}

async function writeSchools(fileName, fileType = "xlsx") {
  let wb = xlsx.readFile(
    path.resolve(__dirname, "../data/" + fileName + "." + fileType)
  );
  // console.log(wb.Sheets["F1,F2,G1"]["F1171"]);
  // console.log(wb.Sheets["F1,F2,G1"]["H1171"]);
  const result = [];
  const sheetName = wb.SheetNames[0];

  for (
    let i = config.data[fileName].rowBegin;
    i <= config.data[fileName].rowEnd;
    i++
  ) {
    // const zipcode =
    //   wb.Sheets[sheetName]["H" + i]?.w || wb.Sheets[sheetName]["H" + i]?.v;

    // if (isNaN(Number(zipcode))) continue;

    const schoolName =
      wb.Sheets[sheetName]["F" + i]?.w || wb.Sheets[sheetName]["F" + i]?.v;
    // result[String(zipcode)] = String(schoolName);

    isValid(config.data[fileName].NGWord, schoolName) &&
      result.push(String(schoolName));
  }

  return await fs.writeFile(
    path.resolve(
      config.writeFileDirName + "/",
      config.data[fileName].writeFileName
    ),
    JSON.stringify(result)
  );
}

async function writePrefectureAndCities(fileName, fileType = "xlsx") {
  let wb = xlsx.readFile(
    path.resolve(__dirname, "../data/" + fileName + "." + fileType)
  );
  // console.log(wb.Sheets["F1,F2,G1"]["F1171"]);
  // console.log(wb.Sheets["F1,F2,G1"]["H1171"]);

  const result = {};
  const sheetName = wb.SheetNames[0];

  for (
    let i = config.data[fileName].rowBegin;
    i <= config.data[fileName].rowEnd;
    i++
  ) {
    const prefecture =
      wb.Sheets[sheetName]["B" + i]?.w || wb.Sheets[sheetName]["B" + i]?.v;
    const city =
      wb.Sheets[sheetName]["C" + i]?.w || wb.Sheets[sheetName]["C" + i]?.v;

    if (!city) {
      result[prefecture] = [];
    } else {
      result[prefecture].push(city);
    }
  }

  return await fs.writeFile(
    path.resolve(
      config.writeFileDirName + "/",
      config.data[fileName].writeFileName
    ),
    JSON.stringify(result)
  );
}

module.exports = { writeSchools, writePrefectureAndCities };
