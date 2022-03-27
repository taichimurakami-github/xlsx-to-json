const { writePrefectureAndCities, writeSchools } = require("./functions");

async function main() {
  await writeSchools("schools_01");

  await writeSchools("schools_02");

  await writePrefectureAndCities("prefecture_and_cities");
}

main();
