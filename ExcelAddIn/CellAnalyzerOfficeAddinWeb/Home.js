(function () {
  "use strict";
  // The initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
    });
  };
})();

function showUnicode() {
  Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    range.load("values");
    return context.sync(range).then(function (range) {
        const url = "https://localhost:44342/api/analyzeunicode?value=" + range.values[0][0];
      $.ajax({
        type: "GET",
        url: url,
        success: function (data) {
          let htmlData = data.replace(/\r\n/g, '<br>');
          $("#txtResult").html(htmlData);
        },
        error: function (data) {
            $("#txtResult").html("error occurred in ajax call.");
        }
      });
    });
  });
}