async function registerClickHandler() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets;
    //context.workbook.onSelectionChanged
    //sheet.onSingleClicked.add((event) => {
    context.workbook.onSelectionChanged.add((event) => {
      return Excel.run(async (context) => {
        //console.log("onSelectionChanged " + event.workbook);
        let selectedtime = await getdetailtime("selection changed : ");
        await recodesheetrange(recodeselectionjsonname,"C1",selectedtime);
        await recoderange(recodeselectionjsonname, "C1", selectedtime);
        
        /*
        //console.log(
          `Click detected at ${event.address} (pixel offset from upper-left cell corner: ${event.offsetX}, ${event.offsetY})`
        );
        */
        //await getcommand();
        return context.sync();
      });
    });

    //console.log("The worksheet click handler is registered.");

    await context.sync();
  });
}
registerClickHandler();
var intervalID = setInterval(resetpreviosindex, 300000);

async function resetpreviosindex(){
  previousindex = 0 ;


}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

const recodesheetname = "recode";
const recodejsonname = "recodejson";
const recodeselectionjsonname = "recodeselectionjson";

var recoderangejson = {
  time: "",
  sheetname: "",
  sheetaddress: "",
  address: "",
  id: ""
};

var recoderangejsonarray = {
  array: []
};

var recodesheetjson = {
  time: "",
  sheetname: "",
  sheetaddress: "",
  address: "",
  id: ""
};

var recodesheetjsonarray = {
  sheetname: "",
  array: []
};
var recodesheetjsonarraycollection = {
  array: []
};

var recodejson = {
  recoderangejsonarray: recoderangejsonarray,
  recodesheetjsonarraycollection: recodesheetjsonarraycollection
};

async function readrecoderange(jsonname,address,id) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(recodesheetname);

    await context.sync();
    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(recodesheetname);
    }
    dataSheet.position = 0;
    let searchRange = dataSheet.getUsedRange();
    let foundRange = searchRange.findOrNullObject(jsonname, {
      completeMatch: true, // Match the whole cell value.
      matchCase: false, // Don't match case.
      searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address,values");
    //let activesheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();

    //console.log("after foundRange ");
    let recodejsonaddress = "";
    let recodejsonva = "";
    if (foundRange.isNullObject) {
      dataSheet.getRange(address).values = [[jsonname]];
      recodejsonaddress = address;
      recodejsonva = "";
      //console.log(`don't find ${jsonname}`);
    } else {
      recodejsonaddress = await get_r_by_address(foundRange.address.toString());
      recodejsonva = foundRange.values[0][0];
      //recodejsonaddress = foundRange.address.toString();
    }
    //console.log(`recodejsonaddress ${recodejsonaddress} ${recodejsonva}`);
    let recodejsonrange = dataSheet.getRange(recodejsonaddress);
    let recodejsonvaluerange = recodejsonrange.getOffsetRange(1, 0);
    recodejsonvaluerange.load("values,valueTypes");

    await context.sync();

    var recodejsonvalue = recodejsonvaluerange.values[0][0];
    //console.log(`recodejsonvalue ${recodejsonvalue}`);
    if (recodejsonvaluerange.valueTypes[0][0] === Excel.RangeValueType.empty) {
      let codejson = recodejson;
      recodejsonvalue = JSON.stringify(codejson);
      //codejson.recoderangejsonarray.array.push(selectedinfomation);
    }
    //console.log(`object parse `);

    //need json.parse or will error
    recodejsonvalue = Object.assign(JSON.parse(JSON.stringify(recodejson)), JSON.parse(recodejsonvalue));
    //console.log(`recodejsonvalue  ${JSON.stringify(recodejsonvalue)}`);
    let recoderangevalue_new = Array.from(
      recodejsonvalue.recoderangejsonarray.array,
      (element) => (element = Object.assign(JSON.parse(JSON.stringify(recoderangejson)), element))
      //(element) => //console.log("element " + JSON.stringify(element))
    );

    //console.log("recoderangevalue_new " + JSON.stringify(recoderangevalue_new));
    //return;
    let codeinfomation = recoderangevalue_new.find((element) => element.id == id);
    if (codeinfomation != undefined) {
      let recodeSheet = context.workbook.worksheets.getItemOrNullObject(codeinfomation.sheetname);

      await context.sync();
      if (!recodeSheet.isNullObject) {
        recodeSheet.activate();
        recodeSheet.getRange(codeinfomation.sheetaddress).select();
        //console.log(`selected recode sheetrange ${codeinfomation.address}`);
      }
    }

    await context.sync();
  });
}

async function recoderange(jsonname,address,id) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(recodesheetname);

    await context.sync();

    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(recodesheetname);
    }
    dataSheet.position = 0;
    let searchRange = dataSheet.getUsedRange();
    let foundRange = searchRange.findOrNullObject(jsonname, {
      completeMatch: true, // Match the whole cell value.
      matchCase: false, // Don't match case.
      searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address,values");
    //let activesheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();

    //console.log("after foundRange ");
    let recodejsonaddress = "";
    let recodejsonva = "";
    if (foundRange.isNullObject) {
      dataSheet.getRange(address).values = [[jsonname]];
      recodejsonaddress = address;
      recodejsonva = "";
      //console.log(`don't find ${recodejsonname}`);
    } else {
      recodejsonaddress = await get_r_by_address(foundRange.address.toString());
      recodejsonva = foundRange.values[0][0];
      //recodejsonaddress = foundRange.address.toString();
    }
    //console.log(`recodejsonaddress ${recodejsonaddress} ${recodejsonva}`);
    let recodejsonrange = dataSheet.getRange(recodejsonaddress);
    let recodejsonvaluerange = recodejsonrange.getOffsetRange(1, 0);
    recodejsonvaluerange.load("values,valueTypes");

    let range = context.workbook.getSelectedRange();

    range.load("values,address");

    await context.sync();

    var recodejsonvalue = recodejsonvaluerange.values[0][0];
    //console.log(`recodejsonvalue ${recodejsonvalue}`);
    if (recodejsonvaluerange.valueTypes[0][0] === Excel.RangeValueType.empty) {
      let codejson = recodejson;
      recodejsonvalue = JSON.stringify(codejson);
      //codejson.recoderangejsonarray.array.push(selectedinfomation);
    }
    //console.log(`object parse `);

    //need json.parse or will error
    recodejsonvalue = Object.assign(JSON.parse(JSON.stringify(recodejson)), JSON.parse(recodejsonvalue));
    //console.log(`recodejsonvalue  ${JSON.stringify(recodejsonvalue)}`);
    let recoderangevalue_new = Array.from(
      recodejsonvalue.recoderangejsonarray.array,
      (element) => (element = Object.assign(JSON.parse(JSON.stringify(recoderangejson)), element))
      //(element) => //console.log("element " + JSON.stringify(element))
    );

    //console.log("recoderangevalue_new " + JSON.stringify(recoderangevalue_new));
    //return;
    recoderangevalue_new = recoderangevalue_new.filter((obj) => obj.id !== id);

    var currentdate = new Date();
    var datetime =
      "Last recode: " +
      currentdate.getDate() +
      "/" +
      (currentdate.getMonth() + 1) +
      "/" +
      currentdate.getFullYear() +
      " @ " +
      currentdate.getHours() +
      ":" +
      currentdate.getMinutes() +
      ":" +
      currentdate.getSeconds();
    const selectedinfomation = recoderangejson;
    selectedinfomation.time = datetime;
    selectedinfomation.address = range.address;
    selectedinfomation.sheetname = await get_s_n_by_address(range.address);
    selectedinfomation.sheetaddress = await get_r_by_address(range.address);
    selectedinfomation.id = id;
    //console.log(`selectedinfomation ${JSON.stringify(selectedinfomation)}`);

    recoderangevalue_new.push(selectedinfomation);

    //console.log(`JSON.stringify(recoderangevalue_new.push) ${JSON.stringify(recoderangevalue_new)}`);

    recodejsonvalue.recoderangejsonarray.array = recoderangevalue_new;
    recodejsonvaluerange.values = [[JSON.stringify(recodejsonvalue)]];
    //console.log(`JSON.stringify(recodejsonvalue) ${JSON.stringify(recodejsonvalue)}`);
    //console.log(`selectedinfomation ${JSON.stringify(selectedinfomation)}`);
    await context.sync();
  });
}

async function readsheetrange(jsonname,address,id) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(recodesheetname);

    await context.sync();

    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(recodesheetname);
    }
    dataSheet.position = 0;
    let searchRange = dataSheet.getUsedRange();
    let foundRange = searchRange.findOrNullObject(jsonname, {
      completeMatch: true, // Match the whole cell value.
      matchCase: false, // Don't match case.
      searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address,values");
    //let activesheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();

    //console.log("after foundRange ");
    let recodejsonaddress = "";
    let recodejsonva = "";
    if (foundRange.isNullObject) {
      dataSheet.getRange(address).values = [[jsonname]];
      recodejsonaddress = address;
      recodejsonva = "";
      //console.log(`don't find ${recodejsonname}`);
    } else {
      recodejsonaddress = await get_r_by_address(foundRange.address.toString());
      recodejsonva = foundRange.values[0][0];
      //recodejsonaddress = foundRange.address.toString();
    }
    //console.log(`recodejsonaddress ${recodejsonaddress} ${recodejsonva}`);
    let recodejsonrange = dataSheet.getRange(recodejsonaddress);
    let recodejsonvaluerange = recodejsonrange.getOffsetRange(1, 0);
    recodejsonvaluerange.load("values,valueTypes");

    let range = context.workbook.getSelectedRange();

    range.load("values,address");

    await context.sync();
    let currentsheetname = await get_s_n_by_address(range.address);

    var recodejsonvalue = recodejsonvaluerange.values[0][0];
    //console.log(`recodejsonvalue ${recodejsonvalue}`);
    if (recodejsonvaluerange.valueTypes[0][0] === Excel.RangeValueType.empty) {
      let codejson = JSON.parse(JSON.stringify(recodejson));
      recodejsonvalue = JSON.stringify(codejson);
      //codejson.recoderangejsonarray.array.push(selectedinfomation);
    }
    //console.log(`object parse `);

    //need json.parse or will error
    recodejsonvalue = Object.assign(JSON.parse(JSON.stringify(recodejson)), JSON.parse(recodejsonvalue));
    //console.log(`recodejsonvalue  ${JSON.stringify(recodejsonvalue)}`);

    let recoderangevalue_pre = Array.from(
      recodejsonvalue.recodesheetjsonarraycollection.array,
      (element) => (element = Object.assign(JSON.parse(JSON.stringify(recodesheetjsonarray)), element))
      //(element) => //console.log("element " + JSON.stringify(element))
    );

    let recodeinfomation = recoderangevalue_pre.find((obj) => obj.sheetname == currentsheetname);
    if (recodeinfomation != undefined) {
      let recoderangevalue_new = Array.from(
        recodeinfomation.array,
        (element) => (element = Object.assign(JSON.parse(JSON.stringify(recodesheetjson)), element))
        //(element) => //console.log("element " + JSON.stringify(element))
      );

      //console.log("recodesheetrangevalue_new " + JSON.stringify(recoderangevalue_new));
      //return;

      let codeinfomation = recoderangevalue_new.find((element) => element.id == id);
      if (codeinfomation != undefined) {
        let recodeSheet = context.workbook.worksheets.getItemOrNullObject(codeinfomation.sheetname);

        await context.sync();
        if (!recodeSheet.isNullObject) {
          recodeSheet.activate();
          recodeSheet.getRange(codeinfomation.sheetaddress).select();
          //console.log(`selected recode sheetrange ${codeinfomation.address}`);
        }
      }

      await context.sync();
    }
  });
}

let recoderangevalue_pre_global = recodesheetjsonarraycollection.array;
async function recodesheetrangeold(id) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(recodesheetname);

    await context.sync();

    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(recodesheetname);
    }
    dataSheet.position = 0;
    let searchRange = dataSheet.getUsedRange();
    let foundRange = searchRange.findOrNullObject(recodejsonname, {
      completeMatch: true, // Match the whole cell value.
      matchCase: false, // Don't match case.
      searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address,values");
    //let activesheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();

    //console.log("after foundRange ");
    let recodejsonaddress = "";
    let recodejsonva = "";
    if (foundRange.isNullObject) {
      dataSheet.getRange("A1").values = [[recodejsonname]];
      recodejsonaddress = "A1";
      recodejsonva = "";
      //console.log(`don't find ${recodejsonname}`);
    } else {
      recodejsonaddress = await get_r_by_address(foundRange.address.toString());
      recodejsonva = foundRange.values[0][0];
      //recodejsonaddress = foundRange.address.toString();
    }
    //console.log(`recodejsonaddress ${recodejsonaddress} ${recodejsonva}`);
    let recodejsonrange = dataSheet.getRange(recodejsonaddress);
    let recodejsonvaluerange = recodejsonrange.getOffsetRange(1, 0);
    recodejsonvaluerange.load("values,valueTypes");

    let range = context.workbook.getSelectedRange();

    range.load("values,address");

    await context.sync();
    let currentsheetname = await get_s_n_by_address(range.address);

    let recodejson_now = recodejson;
    let recodesheetjsonarray_now = recodesheetjsonarray;
    let recodesheetjson_now = recodesheetjson;

    var recodejsonvalue = recodejsonvaluerange.values[0][0];
    //console.log(`recodejsonvalue ${recodejsonvalue}`);
    if (recodejsonvaluerange.valueTypes[0][0] === Excel.RangeValueType.empty) {
      let codejson = recodejson_now;
      recodejsonvalue = JSON.stringify(codejson);
      //codejson.recoderangejsonarray.array.push(selectedinfomation);
    }
    //console.log(`object parse `);

    //need json.parse or will error

    recodejsonvalue = Object.assign(recodejson_now, JSON.parse(recodejsonvalue));
    //console.log(`recodejsonvalue  ${JSON.stringify(recodejsonvalue)}`);
    let recoderangevalue_pre = Array.from(
      recodejsonvalue.recodesheetjsonarraycollection.array,
      (element) => (element = Object.assign(recodesheetjsonarray_now, element))
      //(element) => //console.log("element " + JSON.stringify(element))
    );

    let recodeinfomation = recoderangevalue_pre.find((obj) => obj.sheetname == currentsheetname);
    //console.log(`recoderangevalue_pre  origin : ${JSON.stringify(recoderangevalue_pre)}`);
    if (recodeinfomation != undefined) {
      let recoderangevalue_new = Array.from(
        recodeinfomation.array,
        (element) => (element = Object.assign(recodesheetjson_now, element))
        //(element) => //console.log("element " + JSON.stringify(element))
      );

      //console.log("recodesheetrangevalue_new " + JSON.stringify(recoderangevalue_new));
      //return;
      recoderangevalue_new = recoderangevalue_new.filter((obj) => obj.id !== id);
      recoderangevalue_pre = recoderangevalue_pre.filter((obj) => obj.sheetname !== currentsheetname);

      var currentdate = new Date();
      var datetime =
        "Last recode: " +
        currentdate.getDate() +
        "/" +
        (currentdate.getMonth() + 1) +
        "/" +
        currentdate.getFullYear() +
        " @ " +
        currentdate.getHours() +
        ":" +
        currentdate.getMinutes() +
        ":" +
        currentdate.getSeconds();
      const selectedinfomation = recodesheetjson;
      selectedinfomation.time = datetime;
      selectedinfomation.address = range.address;
      selectedinfomation.sheetname = await get_s_n_by_address(range.address);
      selectedinfomation.sheetaddress = await get_r_by_address(range.address);
      selectedinfomation.id = id;

      //console.log(`selectedinfomation ${JSON.stringify(selectedinfomation)}`);

      recoderangevalue_new.push(selectedinfomation);

      //console.log(`JSON.stringify(recoderangevalue_new.push) ${JSON.stringify(recoderangevalue_new)}`);
      recodeinfomation.array = recoderangevalue_new;
      recodeinfomation.sheetname = currentsheetname;

      recoderangevalue_pre.push(recodeinfomation);

      recodejsonvalue.recodesheetjsonarraycollection.array = recoderangevalue_pre;
      recodejsonvaluerange.values = [[JSON.stringify(recodejsonvalue)]];
      //console.log(`JSON.stringify(recodesheetrange jsonvalue) ${JSON.stringify(recodejsonvalue)}`);
      //console.log(`selectedinfomation ${JSON.stringify(selectedinfomation)}`);
      await context.sync();
    } else {
      var currentdate = new Date();
      var datetime =
        "Last recode: " +
        currentdate.getDate() +
        "/" +
        (currentdate.getMonth() + 1) +
        "/" +
        currentdate.getFullYear() +
        " @ " +
        currentdate.getHours() +
        ":" +
        currentdate.getMinutes() +
        ":" +
        currentdate.getSeconds();
      let selectedinfomation = recodesheetjson;
      selectedinfomation.time = datetime;
      selectedinfomation.address = range.address;
      selectedinfomation.sheetname = await get_s_n_by_address(range.address);
      selectedinfomation.sheetaddress = await get_r_by_address(range.address);
      selectedinfomation.id = id;
      //console.log(`recode sheetrange selectedinfomation ${JSON.stringify(selectedinfomation)}`);
      //recoderangevalue_pre = recoderangevalue_pre.filter((obj) => obj.sheetname !== currentsheetname);
      //console.log(`recoderangevalue_pre add new sheet before []: ${JSON.stringify(recoderangevalue_pre)}`);
      let recoderangevalue_new = [];

      recoderangevalue_new.push(selectedinfomation);
      let recodesheetjsonarray_now = recodesheetjsonarray;

      let recodeinfomation_new = recodesheetjsonarray_now;
      //import !!!! need to store recoderangevalue_pre value ,before after code ,else recoderangevalue_pre value will change ,don't no why 2024/03/25 16:57:42
      recoderangevalue_pre_global = recoderangevalue_pre;
      let recoderangevalue_pre_new = recoderangevalue_pre;

      recodeinfomation_new.array = recoderangevalue_new;
      /*
      console.log(
        `recoderangevalue_pre_global before recodeinfomation_new.sheetname${JSON.stringify(
          recoderangevalue_pre_global  )}` );
      recodeinfomation_new.sheetname = currentsheetname;

      console.log(
        `recoderangevalue_pre_global aafter recodeinfomation_new.sheetname []${JSON.stringify(
          recoderangevalue_pre_global
        )}`
      );
*/
      recoderangevalue_pre_global.push(recodeinfomation_new);
      //console.log(`recoderangevalue_pre_global : ${JSON.stringify(recoderangevalue_pre_global)}`);
      //console.log(`recodeinfomation_new : ${JSON.stringify(recodeinfomation_new)}`);

      //recoderangevalue_pre_global[recoderangevalue_pre_global.length -1 ].sheetname = currentsheetname;

      recodejsonvalue.recodesheetjsonarraycollection.array = recoderangevalue_pre_global;
      recodejsonvaluerange.values = [[JSON.stringify(recodejsonvalue)]];
      //console.log(`JSON.stringify(recodesheetrange jsonvalu) new sheet ${JSON.stringify(recodejsonvalue)}`);
      //console.log(`recode sheet  selectedinfomation ${JSON.stringify(selectedinfomation)}`);
      await context.sync();
    }
  });
}
async function recodesheetrange(jsonname,address,id) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(recodesheetname);

    await context.sync();

    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(recodesheetname);
    }
    dataSheet.position = 0;
    let searchRange = dataSheet.getUsedRange();
    let foundRange = searchRange.findOrNullObject(jsonname, {
      completeMatch: true, // Match the whole cell value.
      matchCase: false, // Don't match case.
      searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address,values");
    //let activesheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();

    //console.log("after foundRange ");
    let recodejsonaddress = "";
    let recodejsonva = "";
    if (foundRange.isNullObject) {
      dataSheet.getRange(address).values = [[jsonname]];
      recodejsonaddress = address;
      recodejsonva = "";
      //console.log(`don't find ${recodejsonname}`);
    } else {
      recodejsonaddress = await get_r_by_address(foundRange.address.toString());
      recodejsonva = foundRange.values[0][0];
      //recodejsonaddress = foundRange.address.toString();
    }
    //console.log(`recodejsonaddress ${recodejsonaddress} ${recodejsonva}`);
    let recodejsonrange = dataSheet.getRange(recodejsonaddress);
    let recodejsonvaluerange = recodejsonrange.getOffsetRange(1, 0);
    recodejsonvaluerange.load("values,valueTypes");

    await context.sync();

    let selectedinfomation = await getrecodesheetinfomation(id);

    let currentsheetname = selectedinfomation.sheetname;

    var recodejsonvalue = recodejsonvaluerange.values[0][0];
    //console.log(`recodejsonvalue ${recodejsonvalue}`);
    if (recodejsonvaluerange.valueTypes[0][0] === Excel.RangeValueType.empty) {
      let codejson = JSON.parse(JSON.stringify(recodejson));
      recodejsonvalue = JSON.stringify(codejson);
      //codejson.recoderangejsonarray.array.push(selectedinfomation);
    }
    //console.log(`object parse `);

    //need json.parse or will error

    recodejsonvalue = Object.assign(JSON.parse(JSON.stringify(recodejson)), JSON.parse(recodejsonvalue));
    //console.log(`recodejsonvalue  ${JSON.stringify(recodejsonvalue)}`);
    let recoderangevalue_pre = Array.from(
      recodejsonvalue.recodesheetjsonarraycollection.array,
      (element) => (element = Object.assign(JSON.parse(JSON.stringify(recodesheetjsonarray)), element))
      //(element) => //console.log("element " + JSON.stringify(element))
    );

    let recodeinfomation = recoderangevalue_pre.find((obj) => obj.sheetname == currentsheetname);
    //console.log(`recoderangevalue_pre  origin : ${JSON.stringify(recoderangevalue_pre)}`);
    if (recodeinfomation != undefined) {
      let recoderangevalue_new = Array.from(
        recodeinfomation.array,
        (element) => (element = Object.assign(JSON.parse(JSON.stringify(recodesheetjson)), element))
        //(element) => //console.log("element " + JSON.stringify(element))
      );

      //console.log("recodesheetrangevalue_new " + JSON.stringify(recoderangevalue_new));
      //return;
      recoderangevalue_new = recoderangevalue_new.filter((obj) => obj.id !== id);
      recoderangevalue_pre = recoderangevalue_pre.filter((obj) => obj.sheetname !== currentsheetname);

      //console.log(`selectedinfomation ${JSON.stringify(selectedinfomation)}`);

      recoderangevalue_new.push(selectedinfomation);

      //console.log(`JSON.stringify(recoderangevalue_new.push) ${JSON.stringify(recoderangevalue_new)}`);
      recodeinfomation.array = recoderangevalue_new;
      recodeinfomation.sheetname = currentsheetname;

      recoderangevalue_pre.push(recodeinfomation);

      recodejsonvalue.recodesheetjsonarraycollection.array = recoderangevalue_pre;
      recodejsonvaluerange.values = [[JSON.stringify(recodejsonvalue)]];
      //console.log(`JSON.stringify(recodesheetrange jsonvalue) ${JSON.stringify(recodejsonvalue)}`);
      //console.log(`selectedinfomation ${JSON.stringify(selectedinfomation)}`);
      await context.sync();
    } else {
      let recoderangevalue_create = await recodesheetrangeinfomationupdate(recoderangevalue_pre, selectedinfomation);

      //console.log(`recoderangevalue_create : ${JSON.stringify(recoderangevalue_create)}`);

      //recoderangevalue_pre_global[recoderangevalue_pre_global.length -1 ].sheetname = currentsheetname;

      recodejsonvalue.recodesheetjsonarraycollection.array = recoderangevalue_create;
      recodejsonvaluerange.values = [[JSON.stringify(recodejsonvalue)]];
      //console.log(`JSON.stringify(recodesheetrange jsonvalu) new sheet ${JSON.stringify(recodejsonvalue)}`);
      //console.log(`recode sheet  selectedinfomation ${JSON.stringify(selectedinfomation)}`);
      await context.sync();
    }
  });
}

async function getrecodesheetinfomation(id) {
  let recodesheetinfomation = recodesheetjson;
  await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.load("values,address");
    await context.sync();
    let currentsheetname = await get_s_n_by_address(range.address);
    var currentdate = new Date();
    var datetime =
      "Last recode: " +
      currentdate.getDate() +
      "/" +
      (currentdate.getMonth() + 1) +
      "/" +
      currentdate.getFullYear() +
      " @ " +
      currentdate.getHours() +
      ":" +
      currentdate.getMinutes() +
      ":" +
      currentdate.getSeconds();
    let selectedinfomation = recodesheetjson;
    selectedinfomation.time = datetime;
    selectedinfomation.address = range.address;
    selectedinfomation.sheetname = await get_s_n_by_address(range.address);
    selectedinfomation.sheetaddress = await get_r_by_address(range.address);
    selectedinfomation.id = id;

    recodesheetinfomation = selectedinfomation;

    await context.sync();
  });
  return recodesheetinfomation;
}

async function recodesheetrangeinfomationupdate(origin, selectedinfomation) {
  //console.log(`origin before push ${JSON.stringify(origin)}`);
  let recoderangevalue_new = [];
  recoderangevalue_new.push(selectedinfomation);
  //https://stackoverflow.com/questions/55711036/on-pushing-object-to-array-all-values-of-one-specific-varaible-in-the-object-bec#:~:text=a%20free%20Team-,On%20pushing%20object%20to%20array%20all%20values%20of%20one%20specific%20varaible%20in%20the%20object%20becomes%20the%20same,-Ask%20Question aviod   incert object to array  failed !
  let recoderangevalue_pre_new = JSON.parse(JSON.stringify(recodesheetjsonarray));
  recoderangevalue_pre_new.array = recoderangevalue_new;
  recoderangevalue_pre_new.sheetname = selectedinfomation.sheetname;

  origin.push(recoderangevalue_pre_new);
  //console.log(` origin : ${JSON.stringify(origin)}`);
  return origin;
}

async function get_s_n_by_address(address) {
  let str = address.toString();
  let str1 = str.slice(0, str.indexOf("!"));
  return str1;
}
async function get_r_by_address(address) {
  let str = address.toString();
  let str2 = str.slice(str.indexOf("!") + 1);
  return str2;
}

async function getdetailtime(info){
  var currentdate = new Date();
  var datetime =
    info.toString() + 
    " : " +
    currentdate.getDate() +
    "/" +
    (currentdate.getMonth() + 1) +
    "/" +
    currentdate.getFullYear() +
    " @ " +
    currentdate.getHours() +
    ":" +
    currentdate.getMinutes() +
    ":" +
    currentdate.getSeconds() +
    ":" + currentdate.getUTCMilliseconds()

    ;
    return datetime.toString();



}

//2024/03/26 10:29:57 add
let previousindex = 0;
async function readrecodeworkbookselection(jsonname, address, index) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(recodesheetname);

    await context.sync();
    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(recodesheetname);
    }
    dataSheet.position = 0;
    let searchRange = dataSheet.getUsedRange();
    let foundRange = searchRange.findOrNullObject(jsonname, {
      completeMatch: true, // Match the whole cell value.
      matchCase: false, // Don't match case.
      searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address,values");
    //let activesheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();

    //console.log("after foundRange ");
    let recodejsonaddress = "";
    let recodejsonva = "";
    if (foundRange.isNullObject) {
      dataSheet.getRange(address).values = [[jsonname]];
      recodejsonaddress = address;
      recodejsonva = "";
      //console.log(`don't find ${jsonname}`);
    } else {
      recodejsonaddress = await get_r_by_address(foundRange.address.toString());
      recodejsonva = foundRange.values[0][0];
      //recodejsonaddress = foundRange.address.toString();
    }
    //console.log(`recodejsonaddress ${recodejsonaddress} ${recodejsonva}`);
    let recodejsonrange = dataSheet.getRange(recodejsonaddress);
    let recodejsonvaluerange = recodejsonrange.getOffsetRange(1, 0);
    recodejsonvaluerange.load("values,valueTypes");

    await context.sync();

    var recodejsonvalue = recodejsonvaluerange.values[0][0];
    //console.log(`recodejsonvalue ${recodejsonvalue}`);
    if (recodejsonvaluerange.valueTypes[0][0] === Excel.RangeValueType.empty) {
      let codejson = recodejson;
      recodejsonvalue = JSON.stringify(codejson);
      //codejson.recoderangejsonarray.array.push(selectedinfomation);
    }
    //console.log(`object parse `);

    //need json.parse or will error
    recodejsonvalue = Object.assign(JSON.parse(JSON.stringify(recodejson)), JSON.parse(recodejsonvalue));
    //console.log(`recodejsonvalue  ${JSON.stringify(recodejsonvalue)}`);
    let recoderangevalue_new = Array.from(
      recodejsonvalue.recoderangejsonarray.array,
      (element) => (element = Object.assign(JSON.parse(JSON.stringify(recoderangejson)), element))
      //(element) => //console.log("element " + JSON.stringify(element))
    );

    //console.log("recoderangevalue_new " + JSON.stringify(recoderangevalue_new));
    //return;
    let codeinfomation = recoderangevalue_new.find((element, indexo) => indexo == (index == 0 ? recoderangevalue_new.length - 2 : (index - 1)));

    previousindex = (index == 0 ? recoderangevalue_new.length - 2 : (index - 1));

    if (codeinfomation != undefined) {
      let recodeSheet = context.workbook.worksheets.getItemOrNullObject(codeinfomation.sheetname);

      await context.sync();
      if (!recodeSheet.isNullObject) {
        recodeSheet.activate();
        recodeSheet.getRange(codeinfomation.sheetaddress).select();
        //console.log(`selected recode sheetrange ${codeinfomation.address}`);
      }
    }

    await context.sync();
  });
}


async function readsheetselection(jsonname, address, index) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(recodesheetname);

    await context.sync();

    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(recodesheetname);
    }
    dataSheet.position = 0;
    let searchRange = dataSheet.getUsedRange();
    let foundRange = searchRange.findOrNullObject(jsonname, {
      completeMatch: true, // Match the whole cell value.
      matchCase: false, // Don't match case.
      searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address,values");
    //let activesheet = context.workbook.worksheets.getActiveWorksheet();

    await context.sync();

    //console.log("after foundRange ");
    let recodejsonaddress = "";
    let recodejsonva = "";
    if (foundRange.isNullObject) {
      dataSheet.getRange(address).values = [[jsonname]];
      recodejsonaddress = address;
      recodejsonva = "";
      //console.log(`don't find ${recodejsonname}`);
    } else {
      recodejsonaddress = await get_r_by_address(foundRange.address.toString());
      recodejsonva = foundRange.values[0][0];
      //recodejsonaddress = foundRange.address.toString();
    }
    //console.log(`recodejsonaddress ${recodejsonaddress} ${recodejsonva}`);
    let recodejsonrange = dataSheet.getRange(recodejsonaddress);
    let recodejsonvaluerange = recodejsonrange.getOffsetRange(1, 0);
    recodejsonvaluerange.load("values,valueTypes");

    let range = context.workbook.getSelectedRange();

    range.load("values,address");

    await context.sync();
    let currentsheetname = await get_s_n_by_address(range.address);

    var recodejsonvalue = recodejsonvaluerange.values[0][0];
    //console.log(`recodejsonvalue ${recodejsonvalue}`);
    if (recodejsonvaluerange.valueTypes[0][0] === Excel.RangeValueType.empty) {
      let codejson = JSON.parse(JSON.stringify(recodejson));
      recodejsonvalue = JSON.stringify(codejson);
      //codejson.recoderangejsonarray.array.push(selectedinfomation);
    }
    //console.log(`object parse `);

    //need json.parse or will error
    recodejsonvalue = Object.assign(JSON.parse(JSON.stringify(recodejson)), JSON.parse(recodejsonvalue));
    //console.log(`recodejsonvalue  ${JSON.stringify(recodejsonvalue)}`);

    let recoderangevalue_pre = Array.from(
      recodejsonvalue.recodesheetjsonarraycollection.array,
      (element) => (element = Object.assign(JSON.parse(JSON.stringify(recodesheetjsonarray)), element))
      //(element) => //console.log("element " + JSON.stringify(element))
    );

    let recodeinfomation = recoderangevalue_pre.find((obj) => obj.sheetname == currentsheetname);
    if (recodeinfomation != undefined) {
      let recoderangevalue_new = Array.from(
        recodeinfomation.array,
        (element) => (element = Object.assign(JSON.parse(JSON.stringify(recodesheetjson)), element))
        //(element) => //console.log("element " + JSON.stringify(element))
      );

      //console.log("recodesheetrangevalue_new " + JSON.stringify(recoderangevalue_new));
      //return;

      let codeinfomation = recoderangevalue_new.find((element, indexo) => indexo == (index == 0 ? recoderangevalue_new.length - 2 : (index - 1)));
      previousindex = (index == 0 ? recoderangevalue_new.length - 2 : (index - 1));
      if (codeinfomation != undefined) {
        let recodeSheet = context.workbook.worksheets.getItemOrNullObject(codeinfomation.sheetname);

        await context.sync();
        if (!recodeSheet.isNullObject) {
          recodeSheet.activate();
          recodeSheet.getRange(codeinfomation.sheetaddress).select();
          //console.log(`selected recode sheetrange ${codeinfomation.address}`);
        }
      }
      await context.sync();
    }
  });
}