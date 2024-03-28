// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    registerClickHandler();
    registerrecodeClickHandler();
    monitorSheetChanges();

    document.getElementById("connectService").onclick = connectService; // in office-apis-helpers.js
    document.getElementById("selectFilter").onclick = insertFilteredData;
    //new Service
    document.getElementById("connectService_new").onclick = reconnectService_new;
    
    updateRibbon();
    updateTaskPaneUI();
  }
});

async function insertFilteredData() {
  try {
    //Determine which data source the user selected from the radio buttons.
    const radioExcel = document.getElementById("communicationFilter");
    if (radioExcel.checked) {
      generateCustomFunction("Communications");
    } else {
      generateCustomFunction("Groceries");
    }
  } catch (error) {
    console.error(error);
  }
}
//
async function reconnectService_new(){
  try {
    await registerClickHandler();
    await registerrecodeClickHandler();
   
  } catch (error) {
    console.error(error);
  }

  
}
