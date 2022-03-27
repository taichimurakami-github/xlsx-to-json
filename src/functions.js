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

function getSchoolType(schoolName) {
  const store = config.searchWordStore;

  for (const [key, props] of Object.entries(store)) {
    for (const values of props) {
      if (schoolName.indexOf(values) !== -1) return key;
    }
  }

  return "others";
}

async function writeSchools(fileName, fileType = "xlsx") {
  //読み出し準備
  let wb = xlsx.readFile(
    path.resolve(__dirname, "../data/" + fileName + "." + fileType)
  );
  const sheetName = wb.SheetNames[0];
  // console.log(wb.Sheets["F1,F2,G1"]["F1171"]);
  // console.log(wb.Sheets["F1,F2,G1"]["H1171"]);

  //結果データ格納Object定義
  const store = config.searchWordStore;
  const result = {};
  for (const key of Object.keys(store)) result[key] = [];
  result.others = []; //オリジナルな学校名の対応がめんどくさいので、getSchoolTypeでマッチしなかった学校名用result

  //処理開始
  for (
    let i = config.data[fileName].rowBegin;
    i <= config.data[fileName].rowEnd;
    i++
  ) {
    // const zipcode =
    //   wb.Sheets[sheetName]["H" + i]?.w || wb.Sheets[sheetName]["H" + i]?.v;

    // if (isNaN(Number(zipcode))) continue;

    //学校名取り出し
    const schoolName =
      wb.Sheets[sheetName]["F" + i]?.w || wb.Sheets[sheetName]["F" + i]?.v;

    //学校種別取得
    const schoolType = getSchoolType(schoolName);
    // console.log(schoolName, schoolType);

    //結果格納
    result[schoolType].push(String(schoolName));
    // console.log(result[schoolType]);

    // isValid(config.data[fileName].NGWord, schoolName) &&
    //   result[schoolType].push(String(schoolName));
  }

  for (const key of Object.keys(result)) {
    if (result[key].length > 0) {
      const data = JSON.stringify(result[key]);
      fs.writeFile(
        path.resolve(
          config.writeFileDirName + "/", //書き出しディレクトリ
          key + ".json" //ファイル名の例： elementary
          // config.data[fileName].writeFileName
        ),
        data
      );
    }
  }
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
