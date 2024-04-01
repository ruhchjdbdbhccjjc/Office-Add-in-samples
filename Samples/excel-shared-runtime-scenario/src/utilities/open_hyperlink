async function openhyperlink() {
  Excel.run(async (context) => {
    console.debug("trying to  go to SelectedRange hyperlink");
    const range = context.workbook.getSelectedRange().load("hyperlink");
    await context.sync();
    if (range.hyperlink) {
      window.open(range.hyperlink.address);
      console.info(`SelectedRange go to  hyperlink ${range.hyperlink.address}`);
    }
    else{
      console.error("trying to open current SelectedRange hyperlink failed ! SelectedRange dont have hyperlink ")


    }
    //return context.sync().then(() => console.log(range.hyperlink));
  });
}
