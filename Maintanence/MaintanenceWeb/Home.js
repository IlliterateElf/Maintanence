'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            $('#set-color').click(setColor);
        });
    });

    async function setColor() {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getActiveWorksheet();

            //Check AntiMal
            let range = sheet.getRange("D2:D87");
            range.load("values");
            await context.sync();

            let vals = range.values;
            let badRanges =[];
            for (let i = 0; i < vals.length; i++) {
                if (vals[i] != "Enabled") {
                    badRanges.push("D" + String(i + 2));
                }
            }

            //Check Network Attack Defense
            range = sheet.getRange("E2:E87");
            range.load("values");
            await context.sync();

            vals = range.values;
            for (let i = 0; i < vals.length; i++) {
                if (vals[i] != "Enabled") {
                    badRanges.push("E" + String(i + 2));
                }
            }

            //Check Firewall
            range = sheet.getRange("G2:G87");
            range.load("values");
            await context.sync();

            vals = range.values;
            for (let i = 0; i < vals.length; i++) {
                if (vals[i] != "Disabled") {
                    badRanges.push("G" + String(i + 2));
                }
            }

            //Check Content Control
            range = sheet.getRange("H2:H87");
            range.load("values");
            await context.sync();

            vals = range.values;
            for (let i = 0; i < vals.length; i++) {
                if (vals[i] != "Enabled") {
                    badRanges.push("H" + String(i + 2));
                }
            }

            console.log(badRanges);
            //let sheet = context.workbook.worksheets.getActiveWorksheet();

            //let rangeUpdate = sheet.getRange("E2:E90");
            //rangeUpdate.load("values");
            //await context.sync();
            //let dateUpdate = [];
            //for (let i = 0; i < rangeUpdate.values.length; i++) {
            //    let newdate = Date.parse(rangeUpdate.values[i]);
            //    newdate = newdate / (1000 * 60 * 60 * 24 * 30);
            //    dateUpdate.push(newdate);
            //}

            //let rangeOnline = sheet.getRange("H2:H90");
            //rangeOnline.load("values");
            //await context.sync();
            //let dateOnline = [];
            //for (let i = 0; i < rangeOnline.values.length; i++) {
            //    if (rangeOnline.values[i] != "Online") {
            //        let newdate = Date.parse(rangeOnline.values[i]);
            //        newdate = newdate / (1000 * 60 * 60 * 24 * 30);
            //        dateOnline.push(newdate);
            //    }
            //    else {
            //        let newdate = new Date();
            //        newdate = Date.parse(newdate);
            //        newdate = newdate / (1000 * 60 * 60 * 24 * 30);
            //        dateOnline.push(newdate);
            //    }
            //}

            //let badIndexes = [];
            //for (let i = 0; i < dateUpdate.length; i++) {
            //    if (dateOnline[i] - dateUpdate[i] > 1) {
            //        badIndexes.push(i);
            //    }
            //}

            //for (let i in badIndexes) {
            //    let row = String(badIndexes[i] + 2);
            //    let range = sheet.getRange("E" + row);
            //    range.format.fill.color = "red";
            //}
            //let sheet = context.workbook.worksheets.getActiveWorksheet();

            //const rangeF = sheet.getRange("F2:F87");
            //const conditionalFormatF = rangeF.conditionalFormats.add(
            //    Excel.ConditionalFormatType.cellValue
            //);

            //// Set the fill of nonzeros to red.
            //conditionalFormatF.cellValue.format.fill.color = "red";
            //conditionalFormatF.cellValue.rule = { formula1: "=0", operator: "GreaterThan" };

            //const rangeI = sheet.getRange("I2:I87");
            //const conditionalFormatI = rangeI.conditionalFormats.add(
            //    Excel.ConditionalFormatType.cellValue
            //);

            //// Set the fill of nonzero numbers to red.
            //conditionalFormatI.cellValue.format.fill.color = "red";
            //conditionalFormatI.cellValue.rule = { formula1: "=0", operator: "GreaterThan" };

            await context.sync();
        });
    }
})();