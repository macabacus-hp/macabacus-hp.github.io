// import { settings } from "./settingsClass";

import { settings } from "../settingsClass.js";

export const handleSettingsChange = async () => {
  await Excel.run(async context => {
    await context.sync();
    console.log("Settings changed");
    JSON.parse(localStorage.getItem("formatSettings")) === null
      ? settings.reset()
      : settings.set(JSON.parse(localStorage.getItem("formatSettings")));
  });
};

export const handleSelectionChange = async () => {

  
  Office.context.document.settings.set("blueBlackToggle", 0);
  Office.context.document.settings.set("fontColorCycle", 0);
  Office.context.document.settings.set("fillColorCycle", 0);
  Office.context.document.settings.set("borderColorCycle", 0);
  Office.context.document.settings.set("chartColorCycle", 0);
  Office.context.document.settings.set("autoColorCycle", 0);

  Office.context.document.settings.set("general", 0);
  Office.context.document.settings.set("localCurrency", 0);
  Office.context.document.settings.set("foreignCurrency", 0);
  Office.context.document.settings.set("percent", 0);
  Office.context.document.settings.set("multiple", 0);
  Office.context.document.settings.set("date", 0);
  Office.context.document.settings.set("binary", 0);
  Office.context.document.settings.set("ratio", 0);

  Office.context.document.settings.set("topBorder", 0);
  Office.context.document.settings.set("bottomBorder", 0);
  Office.context.document.settings.set("leftBorder", 0);
  Office.context.document.settings.set("rightBorder", 0);
  Office.context.document.settings.set("outsideBorder", 0);
  Office.context.document.settings.set("insideBorder", 0);

  Office.context.document.settings.set("centerAlignment", 0);
  Office.context.document.settings.set("horizontalAlignment", 0);
  Office.context.document.settings.set("verticalAlignment", 0);

  Office.context.document.settings.set("leftIndentCycle", 0);
  Office.context.document.settings.set("rightIndentCycle", 0);

  Office.context.document.settings.set("fontSizeCycle", 0);

  Office.context.document.settings.set("underlineCycle", 0);

  Office.context.document.settings.set("caseCycle", 0);

  Office.context.document.settings.set("listCycle", 0);

  Office.context.document.settings.set("footnotesCycle", 0);
  Office.context.document.settings.saveAsync();
  console.log("Cycles Reset");
};
