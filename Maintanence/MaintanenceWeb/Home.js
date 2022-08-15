﻿'use strict';

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

            let rangeUpdate = sheet.getRange("E2:E90");
            rangeUpdate.load("values");
            await context.sync();
            let dateUpdate = [];
            for (let i = 0; i < rangeUpdate.values.length; i++) {
                let newdate = Date.parse(rangeUpdate.values[i]);
                newdate = newdate / (1000*60*60*24*30);
                dateUpdate.push(newdate);
            }

            let rangeOnline = sheet.getRange("H2:H90");
            rangeOnline.load("values");
            await context.sync();
            let dateOnline = [];
            for (let i = 0; i < rangeOnline.values.length; i++) {
                if (rangeOnline.values[i] != "Online") {
                    let newdate = Date.parse(rangeOnline.values[i]);
                    newdate = newdate / (1000 * 60 * 60 * 24 * 30);
                    dateOnline.push(newdate);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              
                }
                else {
                    let newdate = new Date();
                    newdate = Date.parse(newdate);
                    newdate = newdate / (1000 * 60 * 60 * 24 * 30);
                    dateOnline.push(newdate);
                }
            }

            let badIndexes = [];
            for (let i = 0; i < dateUpdate.length; i++) {
                if (dateOnline[i] - dateUpdate[i] > 1) {
                    badIndexes.push(i);
                }
            }
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
