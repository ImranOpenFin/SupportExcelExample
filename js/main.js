
fin.desktop.main(function () {
  
  var view = {};
  [].slice.call(document.querySelectorAll('[id]')).forEach(element => view[element.id] = element);   

  var displayContainers = new Map([
      [view.noConnectionContainer, { windowHeight: 195 }],
      [view.noWorkbooksContainer, { windowHeight: 195 }],
      [view.workbooksContainer, { windowHeight: 830 }]
  ]);

      function initializeUIEvents() {
        view.launchExcelLink.addEventListener("click", function () {
          connectToExcel();
      });
      
      view.newWorkbookLink.addEventListener("click", function () {
        fin.desktop.Excel.addWorkbook();
        console.log('created new workbook');
    });


      }

    // Excel Helper Functions

    function checkConnectionStatus() {
      fin.desktop.Excel.getConnectionStatus(connected => {
          if (connected) {
              onExcelConnected(fin.desktop.Excel);
          } else {
              setDisplayContainer(view.noConnectionContainer);
          }
      });
  }

    function connectToExcel() {
      console.log('connectToExcel');
      return fin.desktop.Excel.run();
      }

    initializeUIEvents();

    fin.desktop.ExcelService.init()
    .then(checkConnectionStatus)
    .catch(err => console.error(err));

fin.desktop.System.getEnvironmentVariable("userprofile", profilePath => {
    view.openWorkbookPath.value = profilePath + "\\Documents\\";
});



  });
