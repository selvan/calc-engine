## About
Virtual calc/excel engine that runs in mobile (React native), web browser(Javascript) and server (Nodejs)

For virtual wordprocessor engine: https://www.github.com/selvan/wordprocessor-engine

## Building

    npm run gen-npm-pkg

## Using
### Recompute calculation
    let CalcEngine=require("./dist/npm_bundle.js");
    let sh = new CalcEngine.Sheet();
    CalcEngine.ScheduleSheetCommands(sh, "set A1 formula SUM(B1:B10)");
    CalcEngine.ScheduleSheetCommands(sh, "set B1 value n 5");
    CalcEngine.ScheduleSheetCommands(sh, "set B2 value n 25");
    CalcEngine.ScheduleSheetCommands(sh, "set B3 formula B2*2");
    CalcEngine.ScheduleSheetCommands(sh, "set B4 formula B3*5");
    CalcEngine.ScheduleSheetCommands(sh, "set B5 formula B4*3");
    sh.RecalcSheet();
    sh.cells['A1'].datavalue;

### Merge & Unmerge
    CalcEngine.ScheduleSheetCommands(sh, "merge B1:C1");
    CalcEngine.ScheduleSheetCommands(sh, "unmerge B1");

### Border top and bottom
    CalcEngine.ScheduleSheetCommands(sh, 'set D2 bt 1px solid rgb(0,0,0)\nset D2 bb 1px solid rgb(0,0,0)');

### Serialize
    var sh = new CalcEngine.Sheet();
    let snapshot_sheet = sh.CreateSheetSave()
    
### Deserialize
    var sh = new CalcEngine.Sheet();
    sh.ParseSheetSave(snapshot_sheet)

### Cell Style

#### Style commands

    These style commands shall be observed by inserting console.log at CalcEngine.ExecuteSheetCommand method 

- Cell border

        set A1 bt 1px solid rgb(0,0,0)
        set A1 br 1px solid rgb(0,0,0)
        set A1 bb 1px solid rgb(0,0,0)
        set A1 bl 1px solid rgb(0,0,0)

- Font family, style & size

        set A1 font italic bold 9pt arial,helvetica,sans-serif

- Foreground & background color

        set A1 color rgb(0,221,0)
        set A1 bgcolor rgb(0,0,170)

- Padding

        set A1 layout padding:4px 4px 4px 4px;vertical-align:*;

- Format number (this is sample, there are many formats)

        set A1 nontextvalueformat #,##0

- Format text (this is sample, there are many formats)

        set A1 textvalueformat text-plain

- Align center

        set A1 cellformat center

- Align vertical

        set A1 layout padding:4px 4px 4px 4px;vertical-align:middle;

Note:
This virtual calc engine is based on https://ethercalc.org/