const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

let wb = xlsx.readFile(path.resolve(__dirname, "schoolCode.xlsx"));
// console.log(wb.Sheets["F1,F2,G1"]["F1171"]);
// console.log(wb.Sheets["F1,F2,G1"]["H1171"]);

const result = [];
const sheetName = "F1,F2,G1";

for (let i = 2; i <= 1195; i++) {
  // const zipcode =
  //   wb.Sheets[sheetName]["H" + i]?.w || wb.Sheets[sheetName]["H" + i]?.v;

  // if (isNaN(Number(zipcode))) continue;

  const schoolName =
    wb.Sheets[sheetName]["F" + i]?.w || wb.Sheets[sheetName]["F" + i]?.v;
  // result[String(zipcode)] = String(schoolName);

  result.push(String(schoolName));
}

fs.writeFileSync(
  path.resolve("result/", "schoolCode.json"),
  JSON.stringify(result)
);
