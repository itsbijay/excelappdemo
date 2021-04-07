// The initialize function is required for all add-ins.
let excel = undefined;
Office.initialize = function (reason) {
    excel = Office.context.Excel;
}

// Create a function for writing to the status div.
function addHederFooter() {
    excel?.run(function (context) {
        var range = context.workbook.getSelectedRange();
        range.format.fill.color = "yellow";
        range.load("address");
        return context.sync()
            .then(function () {
                console.log("The range address was \"" + range.address + "\".");
            });
    });
}