// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    
}

// Create a function for writing to the status div.
function addHederFooter() {
    return Excel.run(function (context) {
        var range = context.workbook.getSelectedRange();
        range.format.fill.color = "yellow";
        range.load("address");
        return context.sync()
            .then(function () {
                console.log("The range address was \"" + range.address + "\".");
            });
    });
}