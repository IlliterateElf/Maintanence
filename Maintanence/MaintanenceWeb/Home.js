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

            const range = sheet.getRange("F2:F87");
            const conditionalFormat = range.conditionalFormats.add(
                Excel.ConditionalFormatType.cellValue
            );

            // Set the font of negative numbers to red.
            conditionalFormat.cellValue.format.fill.color = "red";
            conditionalFormat.cellValue.rule = { formula1: "=0", operator: "GreaterThan" };

            await context.sync();
        });
    }
})();
