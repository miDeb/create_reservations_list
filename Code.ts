export function createList() {
  /// Pads numbers from 1 to 9 with a leading 0.
  function padIndex(number: number): string {
    if (number < 10) {
      return `0${number}`;
    } else {
      return number.toString();
    }
  }

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet = spreadSheet.getActiveSheet();

  const currentSheetName = currentSheet.getName();

  if (currentSheetName.endsWith(" - Liste")) {
    // This macro should not run on sheets that were generated by this very macro.
    return;
  }

  const newSheetName = currentSheetName + " - Liste";
  let targetSheet = spreadSheet.getSheetByName(newSheetName);
  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = spreadSheet.insertSheet(newSheetName);
    // Differentiate list tabs by coloring them gray.
    targetSheet.setTabColor("gray");
    // Prevent focus switch to newly inserted sheet.
    spreadSheet.setActiveSheet(currentSheet);
  }

  targetSheet.getRange(1, 1).setValue("Alphabetische Liste der Besucher");

  const insertions: string[][] = [];

  for (let sourceRow = 9; sourceRow <= 30; sourceRow++) {
    for (let sourceColumn = 4; sourceColumn <= 18; sourceColumn++) {
      const sourceString = (
        currentSheet.getRange(sourceRow, sourceColumn).getValue() as string
      ).trim();

      if (sourceString !== "") {
        insertions.push([
          sourceString,
          `Reihe - ${padIndex(sourceRow - 8)}`,
          `Nummer - ${sourceColumn - 3}`,
        ]);
      }
    }
  }

  // Sorting is stable, so seats with the same name will remain sorted according to their row and column
  // (as they were inserted by the nested loop above).
  insertions.sort((a, b) =>
    a[0].toLowerCase().localeCompare(b[0].toLowerCase())
  );

  // Add an incrementing index where multiple seats pertain to the same name:

  let previousName: string | undefined;
  let previousNameCount = 0;
  for (let i = 0; i < insertions.length; i++) {
    const name: string = insertions[i][0];
    let index: number | undefined;
    if (name === previousName) {
      // If we previously inserted this name, continue to insert the incremented index.
      index = ++previousNameCount;
    } else if (insertions[i + 1] && insertions[i + 1][0] === name) {
      // If this name occurs first in this cell but also occurs in the next cell, start by inserting index 1.
      previousName = name;
      previousNameCount = 1;
      index = 1;
    }
    if (index) {
      insertions[i][0] = `${name} - ${index}`;
    }
  }

  // API calls are very expensive, so by batching all changes in `insertions` we avoid performance cliffs.
  targetSheet.getRange(2, 1, insertions.length, 3).setValues(insertions);

  // Make sure the first column is big enough.
  targetSheet.autoResizeColumn(1);
}
