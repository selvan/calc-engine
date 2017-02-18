//
/*
// The module of the SocialCalc package with customizable constants, strings, etc.
// This is where most of the common localizations are done.
//
// (c) Copyright 2008, 2009, 2010 Socialtext, Inc.
// All Rights Reserved.
//
// The contents of this file are subject to the Artistic License 2.0; you may not
// use this file except in compliance with the License. You may obtain a copy of 
// the License at http://socialcalc.org/licenses/al-20/.
//
// Some of the other files in the SocialCalc package are licensed under
// different licenses. Please note the licenses of the modules you use.
//
// Code History:
//
// Initially coded by Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
// Based in part on the SocialCalc 1.1.0 code written in Perl.
// The SocialCalc 1.1.0 code was:
//    Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc.
//    All Rights Reserved.
//    Portions (c) Copyright 2007 Socialtext, Inc.
//    All Rights Reserved.
// The Perl SocialCalc started as modifications to the wikiCalc(R) program, version 1.0.
// wikiCalc 1.0 was written by Software Garden, Inc.
// Unless otherwise specified, referring to "SocialCalc" in comments refers to this
// JavaScript version of the code, not the SocialCalc Perl code.
//
*/

var SocialCalc;
if (!SocialCalc) SocialCalc = {};

// *************************************
//
// TO LEARN HOW TO LOCALIZE OR CUSTOMIZE SOCIALCALC, PLEASE READ THIS:
//
// The constants are all properties of the SocialCalc.Constants object.
// They are grouped here by what they are for, which module uses them, etc.
//
// Properties whose names start with "s_" are strings, or arrays of strings,
// that are good candidates for translation from the English.
//
// Other properties relate to visual settings, localization parameters, etc.
//
// These values are not used when SocialCalc modules are first loaded.
// They may be modified before the first use of the routines that use them,
// e.g., before creating SocialCalc objects.
//
// The exceptions are:
//    TooltipOffsetX and TooltipOffsetY, as described with their definitions.
//
// SocialCalc IS NOT DESIGNED FOR USE WITH A TRANSLATION FUNCTION each time a string
// is used. Instead, language translations may be done by modifying this object.
//
// To customize SocialCalc, you may either replace this file with a modified version
// or you can overwrite the values before use. An example would be to
// iterate over all the properties looking for names that start with "s_" and
// use some other mechanism to obtain a localized string and replace the values
// here with those translated values.
//
// There is also a function, SocialCalc.ConstantsSetClasses, that may be used
// to easily switch SocialCalc from using explicit CSS styles for many things
// to using CSS classes. See the function, below, for more information.
//
// *************************************

SocialCalc.Constants = {

//
// Main SocialCalc module, socialcalc-3.js:
//

   //*** Common Constants

   textdatadefaulttype: "t", // This sets the default type for text on reading source file
                             // It should normally be "t"

   //*** Common error messages

   s_BrowserNotSupported: "Browser not supported.", // error thrown if browser can't handle events like IE or Firefox.
   s_InternalError: "Internal SocialCalc error (probably an internal bug): ", // hopefully unlikely, but a test failed

   //*** SocialCalc.ParseSheetSave

   // Errors thrown on unexpected value in save file:

   s_pssUnknownColType: "Unknown col type item",
   s_pssUnknownRowType: "Unknown row type item",
   s_pssUnknownLineType: "Unknown line type",

   //*** SocialCalc.CellFromStringParts

   // Error thrown on unexpected value in save file:

   s_cfspUnknownCellType: "Unknown cell type item",

   //*** SocialCalc.CanonicalizeSheet

   doCanonicalizeSheet: true, // if true, do the canonicalization calculations

   //*** ExecuteSheetCommand

   s_escUnknownSheetCmd: "Unknown sheet command: ",
   s_escUnknownSetCoordCmd: "Unknown set coord command: ",
   s_escUnknownCmd: "Unknown command: ",

   //*** SocialCalc.CheckAndCalcCell

   s_caccCircRef: "Circular reference to ", // circular reference found during recalc

   //*** SocialCalc.RenderContext

   defaultRowNameWidth: "30", // used to set minimum width of the row header column - a string in pixels
   defaultAssumedRowHeight: 15, // used when guessing row heights - number
   defaultCellIDPrefix: "cell_", // if non-null, each cell will render with an ID starting with this

   // Default sheet display values

   defaultCellLayout: "padding:2px 2px 1px 2px;vertical-align:top;",
   defaultCellFontStyle: "normal normal",
   defaultCellFontSize: "small",
   defaultCellFontFamily: "Verdana,Arial,Helvetica,sans-serif",

   defaultPaneDividerWidth: "2", // a string
   defaultPaneDividerHeight: "3", // a string

   defaultGridCSS: "1px solid #C0C0C0;", // used as style to set each border when grid enabled (was #ECECEC)

   defaultCommentClass: "", // class added to cells with non-null comments when grid enabled
   defaultCommentStyle: "background-repeat:no-repeat;background-position:top right;background-image:url(images/sc-commentbg.gif);", // style added to cells with non-null comments when grid enabled
   defaultCommentNoGridClass: "", // class added to cells with non-null comments when grid not enabled
   defaultCommentNoGridStyle: "", // style added to cells with non-null comments when grid not enabled

   defaultReadonlyClass: "", // class added to readonly cells when grid enabled
   defaultReadonlyStyle: "background-repeat:no-repeat;background-position:top right;background-image:url(images/sc-lockbg.gif);", // style added to readonly cells when grid enabled
   defaultReadonlyNoGridClass: "", // class added to readonly cells when grid not enabled
   defaultReadonlyNoGridStyle: "", // style added to readonly cells when grid not enabled
   defaultReadonlyComment: "Locked cell",

   defaultColWidth: "80", // text
   defaultMinimumColWidth: 10, // numeric

   // For each of the following default sheet display values at least one of class and/or style are needed

   defaultHighlightTypeCursorClass: "",
   defaultHighlightTypeCursorStyle: "color:#FFF;backgroundColor:#A6A6A6;",
   defaultHighlightTypeRangeClass: "",
   defaultHighlightTypeRangeStyle: "color:#000;backgroundColor:#E5E5E5;",

   defaultColnameClass: "", // regular column heading letters, needs a cursor property 
   defaultColnameStyle: "font-size:small;text-align:center;color:#FFFFFF;background-color:#808080;cursor:e-resize;",
   defaultSelectedColnameClass: "", // column with selected cell, needs a cursor property 
   defaultSelectedColnameStyle: "font-size:small;text-align:center;color:#FFFFFF;background-color:#404040;cursor:e-resize;",
   defaultRownameClass: "", // regular row heading numbers
   defaultRownameStyle: "font-size:small;text-align:right;color:#FFFFFF;background-color:#808080;direction:rtl;",
   defaultSelectedRownameClass: "", // column with selected cell, needs a cursor property 
   defaultSelectedRownameStyle: "font-size:small;text-align:right;color:#FFFFFF;background-color:#404040;",
   defaultUpperLeftClass: "", // Corner cell in upper left
   defaultUpperLeftStyle: "font-size:small;",
   defaultSkippedCellClass: "", // used if present for spanned cells peeking into a pane (at least one of class/style needed)
   defaultSkippedCellStyle: "font-size:small;background-color:#CCC", // used if present
   defaultPaneDividerClass: "", // used if present for the look of the space between panes (at least one of class/style needed)
   defaultPaneDividerStyle: "font-size:small;background-color:#C0C0C0;padding:0px;", // used if present
   defaultUnhideLeftClass: "",
   defaultUnhideLeftStyle: "float:right;width:9px;height:12px;cursor:pointer;background-image:url(images/sc-unhideleft.gif);padding:0;", // used if present
   defaultUnhideRightClass: "",
   defaultUnhideRightStyle: "float:left;width:9px;height:12px;cursor:pointer;background-image:url(images/sc-unhideright.gif);padding:0;", // used if present
   defaultUnhideTopClass: "",
   defaultUnhideTopStyle: "float:left;position:absolute;bottom:-4px;width:12px;height:9px;cursor:pointer;background-image:url(images/sc-unhidetop.gif);padding:0;",
   defaultUnhideBottomClass: "",
   defaultUnhideBottomStyle: "float:left;width:12px;height:9px;cursor:pointer;background-image:url(images/sc-unhidebottom.gif);padding:0;",

   s_rcMissingSheet: "Render Context must have a sheet object", // unlikely thrown error

   //*** SocialCalc.format_text_for_display

   defaultLinkFormatString: '<span style="font-size:smaller;text-decoration:none !important;background-color:#66B;color:#FFF;">Link</span>', // used for format "text-link"; you could make this an img tag if desired
//   defaultLinkFormatString: '<img src="images/sc-linkout.gif" border="0" alt="Link out" title="Link out">',
   defaultPageLinkFormatString: '<span style="font-size:smaller;text-decoration:none !important;background-color:#66B;color:#FFF;">Page</span>', // used for format "text-link"; you could make this an img tag if desired

   //*** SocialCalc.format_number_for_display

   defaultFormatp: '#,##0.0%',
   defaultFormatc: '[$$]#,##0.00',
   defaultFormatdt: 'd-mmm-yyyy h:mm:ss',
   defaultFormatd: 'd-mmm-yyyy',
   defaultFormatt: '[h]:mm:ss',
   defaultDisplayTRUE: 'TRUE', // how TRUE shows when rendered
   defaultDisplayFALSE: 'FALSE',

//
// SocialCalc Table Editor module, socialcalctableeditor.js:
//

   //*** SocialCalc.TableEditor

   defaultImagePrefix: "images/sc-", // URL prefix for images (e.g., "/images/sc")
   defaultTableEditorIDPrefix: "te_", // if present, many TableEditor elements are assigned IDs with this prefix
   defaultPageUpDnAmount: 15, // number of rows to move cursor on PgUp/PgDn keys (numeric)

   AllowCtrlS: true, // turns on Ctrl-S trapdoor for setting custom numeric formats and commands if true

   //*** SocialCalc.CreateTableEditor

   defaultTableControlThickness: 20, // the short size for the scrollbars, etc. (numeric in pixels)
   cteGriddivClass: "", // if present, the class for the TableEditor griddiv element

   //** SocialCalc.EditorGetStatuslineString -- strings shown on status line

   s_statusline_executing: "Executing...",
   s_statusline_displaying: "Displaying...",
   s_statusline_ordering: "Ordering...",
   s_statusline_calculating: "Calculating...",
   s_statusline_calculatingls: "Calculating... Loading Sheet...",
   s_statusline_doingserverfunc: "doing server function ",
   s_statusline_incell: " in cell ",
   s_statusline_calcstart: "Calculation start...",
   s_statusline_sum: "SUM",
   s_statusline_recalcneeded: '<span style="color:#999;">(Recalc needed)</span>',
   s_statusline_circref: '<span style="color:red;">Circular reference: ',

   //** SocialCalc.InputBoxDisplayCellContents

   s_inputboxdisplaymultilinetext: "[Multi-line text: Click icon on right to edit]",

   //** SocialCalc.InputEcho

   defaultInputEchoClass: "", // if present, the class of the popup inputEcho div
   defaultInputEchoStyle: "filter:alpha(opacity=90);opacity:.9;backgroundColor:#FFD;border:1px solid #884;"+
      "fontSize:small;padding:2px 10px 1px 2px;cursor:default;", // if present, pseudo style
   defaultInputEchoPromptClass: "", // if present, the class of the popup inputEcho div
   defaultInputEchoPromptStyle: "filter:alpha(opacity=90);opacity:.9;backgroundColor:#FFD;"+
      "borderLeft:1px solid #884;borderRight:1px solid #884;borderBottom:1px solid #884;"+
      "fontSize:small;fontStyle:italic;padding:2px 10px 1px 2px;cursor:default;", // if present, pseudo style

   //** SocialCalc.InputEchoText

   ietUnknownFunction: "Unknown function ", // displayed when typing "=unknown("

   //** SocialCalc.CellHandles

   CH_radius1: 29.0, // extent of inner circle within 90px image
   CH_radius2: 41.0, // extent of outer circle within 90px image
   s_CHfillAllTooltip: "Fill Contents and Formats Down/Right", // tooltip for fill all handle
   s_CHfillContentsTooltip: "Fill Contents Only Down/Right", // tooltip for fill formulas handle
   s_CHmovePasteAllTooltip: "Move Contents and Formats", // etc.
   s_CHmovePasteContentsTooltip: "Move Contents Only",
   s_CHmoveInsertAllTooltip: "Slide Contents and Formats within Row/Col",
   s_CHmoveInsertContentsTooltip: "Slide Contents within Row/Col",
   s_CHindicatorOperationLookup: {"Fill": "Fill", "FillC": "Fill Contents",
                                  "Move": "Move", "MoveI": "Slide", 
                                  "MoveC": "Move Contents", "MoveIC": "Slide Contents"}, // short form of operation to follow drag
   s_CHindicatorDirectionLookup: {"Down": " Down", "Right": " Right",
                                  "Horizontal": " Horizontal", "Vertical": " Vertical"}, // direction that modifies operation during drag

   //*** SocialCalc.TableControl

   defaultTCSliderThickness: 9, // length of pane slider (numeric in pixels)
   defaultTCButtonThickness: 20, // length of scroll +/- buttons (numeric in pixels)
   defaultTCThumbThickness: 15, // length of thumb (numeric in pixels)

   //*** SocialCalc.CreateTableControl

   TCmainStyle: "backgroundColor:#EEE;", // if present, pseudo style (text-align is textAlign) for main div of a table control
   TCmainClass: "", // if present, the CSS class of the main div for a table control
   TCendcapStyle: "backgroundColor:#FFF;", // backgroundColor may be used while waiting for image that may not come
   TCendcapClass: "",
   TCpanesliderStyle: "backgroundColor:#CCC;",
   TCpanesliderClass: "",
   s_panesliderTooltiph: "Drag to lock pane vertically", // tooltip for horizontal table control pane slider
   s_panesliderTooltipv: "Drag to lock pane horizontally",
   TClessbuttonStyle: "backgroundColor:#AAA;",
   TClessbuttonClass: "",
   TClessbuttonRepeatWait: 300, // in milliseconds
   TClessbuttonRepeatInterval: 20,//100, // in milliseconds
   TCmorebuttonStyle: "backgroundColor:#AAA;",
   TCmorebuttonClass: "",
   TCmorebuttonRepeatWait: 300, // in milliseconds
   TCmorebuttonRepeatInterval: 20,//100, // in milliseconds
   TCscrollareaStyle: "backgroundColor:#DDD;",
   TCscrollareaClass: "",
   TCscrollareaRepeatWait: 500, // in milliseconds
   TCscrollareaRepeatInterval: 100, // in milliseconds
   TCthumbClass: "",
   TCthumbStyle: "backgroundColor:#CCC;",

   //*** SocialCalc.TCPSDragFunctionStart

   TCPStrackinglineClass: "", // at least one of class/style for pane slider tracking line display in table control
   TCPStrackinglineStyle: "overflow:hidden;position:absolute;zIndex:100;",
                           // if present, pseudo style (text-align is textAlign)
   TCPStrackinglineThickness: "2px", // narrow dimension of trackling line (string with units)


   //*** SocialCalc.TCTDragFunctionStart

   TCTDFSthumbstatusvClass: "", // at least one of class/style for vertical thumb dragging status display in table control
   TCTDFSthumbstatusvStyle: "height:20px;width:auto;border:3px solid #808080;overflow:hidden;"+
                           "backgroundColor:#FFF;fontSize:small;position:absolute;zIndex:100;",
                           // if present, pseudo style (text-align is textAlign)
   TCTDFSthumbstatushClass: "", // at least one of class/style for horizontal thumb dragging status display in table control
   TCTDFSthumbstatushStyle: "height:20px;width:auto;border:1px solid black;padding:2px;"+
                           "backgroundColor:#FFF;fontSize:small;position:absolute;zIndex:100;",
                           // if present, pseudo style (text-align is textAlign)
   TCTDFSthumbstatusrownumClass: "", // at least one of class/style for thumb dragging status display in table control
   TCTDFSthumbstatusrownumStyle: "color:#FFF;background-color:#808080;font-size:small;white-space:nowrap;padding:3px;", // if present, real style
   TCTDFStopOffsetv: 0, // offsets for thumbstatus display while dragging
   TCTDFSleftOffsetv: -80,
   s_TCTDFthumbstatusPrefixv: "Row ", // Text Control Drag Function text before row number
   TCTDFStopOffseth: -30,
   TCTDFSleftOffseth: 0,
   s_TCTDFthumbstatusPrefixh: "Col ", // Text Control Drag Function text before col number

   //*** SocialCalc.TooltipInfo

   // Note: These two values are used to set the TooltipInfo initial values when the code is first read in.
   // Modifying them here after loading has no effect -- you need to modify SocialCalc.TooltipInfo directly
   // to dynamically set them. This is different than most other constants which may be modified until use.

   TooltipOffsetX: 2, // offset in pixels from mouse position (to right on left side of screen, to left on right)
   TooltipOffsetY: 10, // offset in pixels above mouse position for lower edge

   //*** SocialCalc.TooltipDisplay

   TDpopupElementClass: "", // at least one of class/style for tooltip display
   TDpopupElementStyle: "border:1px solid black;padding:1px 2px 2px 2px;textAlign:center;backgroundColor:#FFF;"+
                        "fontSize:7pt;fontFamily:Verdana,Arial,Helvetica,sans-serif;"+
                        "position:absolute;width:auto;zIndex:110;",
                        // if present, pseudo style (text-align is textAlign)


//
// SocialCalc Spreadsheet Control module, socialcalcspreadsheetcontrol.js:
//

   //*** SocialCalc.SpreadsheetControl

   SCToolbarbackground: "background-color:#404040;",
   SCTabbackground: "background-color:#CCC;",
   SCTabselectedCSS: "font-size:small;padding:6px 30px 6px 8px;color:#FFF;background-color:#404040;cursor:default;border-right:1px solid #CCC;",
   SCTabplainCSS: "font-size:small;padding:6px 30px 6px 8px;color:#FFF;background-color:#808080;cursor:default;border-right:1px solid #CCC;",
   SCToolbartext: "font-size:x-small;font-weight:bold;color:#FFF;padding-bottom:4px;",

   SCFormulabarheight: 30, // in pixels, will contain a text input box

   SCStatuslineheight: 20, // in pixels
   SCStatuslineCSS: "font-size:10px;padding:3px 0px;",

   // Constants for default Format tab (settings)
   //
   // *** EVEN THOUGH THESE DON'T START WITH s_: ***
   //
   // These should be carefully checked for localization. Make sure you understand what they do and how they work!
   // The first part of "first:second|first:second|..." is what is displayed and the second is the value to be used.
   // The value is normally not translated -- only the displayed part. The [cancel], [break], etc., are not translated --
   // they are commands to SocialCalc.SettingsControls.PopupListInitialize 

   SCFormatNumberFormats: "[cancel]:|[break]:|%loc!Default!:|[custom]:|%loc!Automatic!:general|%loc!Auto w/ commas!:[,]General|[break]:|"+
            "00:00|000:000|0000:0000|00000:00000|[break]:|%loc!Formula!:formula|%loc!Hidden!:hidden|[newcol]:"+
            "1234:0|1,234:#,##0|1,234.5:#,##0.0|1,234.56:#,##0.00|1,234.567:#,##0.000|1,234.5678:#,##0.0000|"+
            "[break]:|1,234%:#,##0%|1,234.5%:#,##0.0%|1,234.56%:#,##0.00%|"+
            "[newcol]:|$1,234:$#,##0|$1,234.5:$#,##0.0|$1,234.56:$#,##0.00|[break]:|"+
            "(1,234):#,##0_);(#,##0)|(1,234.5):#,##0.0_);(#,##0.0)|(1,234.56):#,##0.00_);(#,##0.00)|[break]:|"+
            "($1,234):#,##0_);($#,##0)|($1,234.5):$#,##0.0_);($#,##0.0)|($1,234.56):$#,##0.00_);($#,##0.00)|"+
            "[newcol]:|1/4/06:m/d/yy|01/04/2006:mm/dd/yyyy|2006-01-04:yyyy-mm-dd|4-Jan-06:d-mmm-yy|04-Jan-2006:dd-mmm-yyyy|January 4, 2006:mmmm d, yyyy|"+
            "[break]:|1\\c23:h:mm|1\\c23 PM:h:mm AM/PM|1\\c23\\c45:h:mm:ss|01\\c23\\c45:hh:mm:ss|26\\c23 (h\\cm):[hh]:mm|69\\c45 (m\\cs):[mm]:ss|69 (s):[ss]|"+
            "[newcol]:|2006-01-04 01\\c23\\c45:yyyy-mm-dd hh:mm:ss|January 4, 2006:mmmm d, yyyy hh:mm:ss|Wed:ddd|Wednesday:dddd|",
   SCFormatTextFormats: "[cancel]:|[break]:|%loc!Default!:|[custom]:|%loc!Automatic!:general|%loc!Plain Text!:text-plain|"+
            "HTML:text-html|%loc!Wikitext!:text-wiki|%loc!Link!:text-link|%loc!Formula!:formula|%loc!Hidden!:hidden|",
   SCFormatPadsizes: "[cancel]:|[break]:|%loc!Default!:|[custom]:|%loc!No padding!:0px|"+
            "[newcol]:|1 pixel:1px|2 pixels:2px|3 pixels:3px|4 pixels:4px|5 pixels:5px|"+
            "6 pixels:6px|7 pixels:7px|8 pixels:8px|[newcol]:|9 pixels:9px|10 pixels:10px|11 pixels:11px|"+
            "12 pixels:12px|13 pixels:13px|14 pixels:14px|16 pixels:16px|"+
            "18 pixels:18px|[newcol]:|20 pixels:20px|22 pixels:22px|24 pixels:24px|28 pixels:28px|36 pixels:36px|",
   SCFormatFontsizes: "[cancel]:|[break]:|%loc!Default!:|[custom]:|X-Small:x-small|Small:small|Medium:medium|Large:large|X-Large:x-large|"+
                  "[newcol]:|6pt:6pt|7pt:7pt|8pt:8pt|9pt:9pt|10pt:10pt|11pt:11pt|12pt:12pt|14pt:14pt|16pt:16pt|"+
                  "[newcol]:|18pt:18pt|20pt:20pt|22pt:22pt|24pt:24pt|28pt:28pt|36pt:36pt|48pt:48pt|72pt:72pt|"+
                  "[newcol]:|8 pixels:8px|9 pixels:9px|10 pixels:10px|11 pixels:11px|"+
                  "12 pixels:12px|13 pixels:13px|14 pixels:14px|[newcol]:|16 pixels:16px|"+
                  "18 pixels:18px|20 pixels:20px|22 pixels:22px|24 pixels:24px|28 pixels:28px|36 pixels:36px|",
   SCFormatFontfamilies: "[cancel]:|[break]:|%loc!Default!:|[custom]:|Verdana:Verdana,Arial,Helvetica,sans-serif|"+
                  "Arial:arial,helvetica,sans-serif|Courier:'Courier New',Courier,monospace|",
   SCFormatFontlook: "[cancel]:|[break]:|%loc!Default!:|%loc!Normal!:normal normal|%loc!Bold!:normal bold|%loc!Italic!:italic normal|"+
                  "%loc!Bold Italic!:italic bold",
   SCFormatTextAlignhoriz:  "[cancel]:|[break]:|%loc!Default!:|%loc!Left!:left|%loc!Center!:center|%loc!Right!:right|",
   SCFormatNumberAlignhoriz:  "[cancel]:|[break]:|%loc!Default!:|%loc!Left!:left|%loc!Center!:center|%loc!Right!:right|",
   SCFormatAlignVertical: "[cancel]:|[break]:|%loc!Default!:|%loc!Top!:top|%loc!Middle!:middle|%loc!Bottom!:bottom|",
   SCFormatColwidth: "[cancel]:|[break]:|%loc!Default!:|[custom]:|[newcol]:|"+
                  "20 pixels:20|40:40|60:60|80:80|100:100|120:120|140:140|160:160|"+
                  "[newcol]:|180 pixels:180|200:200|220:220|240:240|260:260|280:280|300:300|",
   SCFormatRecalc: "[cancel]:|[break]:|%loc!Auto!:|%loc!Manual!:off|",
   SCFormatUserMaxCol: "[cancel]:|[break]:|%loc!Default!:|[custom]:|[newcol]:|"+
                  "Unlimited:0|10:10|20:20|30:30|40:40|50:50|60:60|80:80|100:100|",
   SCFormatUserMaxRow: "[cancel]:|[break]:|%loc!Default!:|[custom]:|[newcol]:|"+
                  "Unlimited:0|10:10|20:20|30:30|40:40|50:50|60:60|80:80|100:100|",

   //*** SocialCalc.InitializeSpreadsheetControl

   ISCButtonBorderNormal: "#404040",
   ISCButtonBorderHover: "#999",
   ISCButtonBorderDown: "#FFF",
   ISCButtonDownBackground: "#888",

   //*** SocialCalc.SettingsControls.PopupListInitialize

   s_PopupListCancel: "[Cancel]",
   s_PopupListCustom: "Custom",

   // ***
   //
   // s_loc_ constants accessed by SocialCalc.LocalizeString and SocialCalc.LocalizeSubstrings
   //
   // Used extensively by socialcalcspreadsheetcontrol.js
   //
   // ***

   s_loc_align_center: "Align Center",
   s_loc_align_left: "Align Left",
   s_loc_align_right: "Align Right",
   s_loc_alignment: "Alignment",
   s_loc_audit: "Audit",
   s_loc_audit_trail_this_session: "Audit Trail This Session",
   s_loc_auto: "Auto",
   s_loc_auto_sum: "Auto Sum",
   s_loc_auto_wX_commas: "Auto w/ commas",
   s_loc_automatic: "Automatic",
   s_loc_background: "Background",
   s_loc_bold: "Bold",
   s_loc_bold_XampX_italics: "Bold &amp; Italics",
   s_loc_bold_italic: "Bold Italic",
   s_loc_borders: "Borders",
   s_loc_borders_off: "Borders Off",
   s_loc_borders_on: "Borders On",
   s_loc_bottom: "Bottom",
   s_loc_bottom_border: "Bottom Border",
   s_loc_cell_settings: "CELL SETTINGS",
   s_loc_csv_format: "CSV format",
   s_loc_cancel: "Cancel",
   s_loc_category: "Category",
   s_loc_center: "Center",
   s_loc_clear: "Clear",
   s_loc_clear_socialcalc_clipboard: "Clear SocialCalc Clipboard",
   s_loc_clipboard: "Clipboard",
   s_loc_color: "Color",
   s_loc_column_: "Column ",
   s_loc_comment: "Comment",
   s_loc_copy: "Copy",
   s_loc_custom: "Custom",
   s_loc_cut: "Cut",
   s_loc_default: "Default",
   s_loc_default_alignment: "Default Alignment",
   s_loc_default_column_width: "Default Column Width",
   s_loc_default_font: "Default Font",
   s_loc_default_format: "Default Format",
   s_loc_default_padding: "Default Padding",
   s_loc_delete: "Delete",
   s_loc_delete_column: "Delete Column",
   s_loc_delete_contents: "Delete Contents",
   s_loc_delete_row: "Delete Row",
   s_loc_description: "Description",
   s_loc_display_clipboard_in: "Display Clipboard in",
   s_loc_down: "Down",
   s_loc_edit: "Edit",
   s_loc_existing_names: "Existing Names",
   s_loc_family: "Family",
   s_loc_fill_down: "Fill Down",
   s_loc_fill_right: "Fill Right",
   s_loc_font: "Font",
   s_loc_format: "Format",
   s_loc_formula: "Formula",
   s_loc_function_list: "Function List",
   s_loc_functions: "Functions",
   s_loc_grid: "Grid",
   s_loc_hidden: "Hidden",
   s_loc_horizontal: "Horizontal",
   s_loc_insert_column: "Insert Column",
   s_loc_insert_row: "Insert Row",
   s_loc_italic: "Italic",
   s_loc_last_sort: "Last Sort",
   s_loc_left: "Left",
   s_loc_left_border: "Left Border",
   s_loc_link: "Link",
   s_loc_link_input_box: "Link Input Box",
   s_loc_list: "List",
   s_loc_load_socialcalc_clipboard_with_this: "Load SocialCalc Clipboard With This",
   s_loc_major_sort: "Major Sort",
   s_loc_manual: "Manual",
   s_loc_merge_cells: "Merge Cells",
   s_loc_middle: "Middle",
   s_loc_minor_sort: "Minor Sort",
   s_loc_move_insert: "Move Insert",
   s_loc_move_paste: "Move Paste",
   s_loc_multiXline_input_box: "Multi-line Input Box",
   s_loc_name: "Name",
   s_loc_names: "Names",
   s_loc_no_padding: "No padding",
   s_loc_normal: "Normal",
   s_loc_number: "Number",
   s_loc_number_horizontal: "Number Horizontal",
   s_loc_ok: "OK",
   s_loc_padding: "Padding",
   s_loc_page_name: "Page Name",
   s_loc_paste: "Paste",
   s_loc_paste_formats: "Paste Formats",
   s_loc_plain_text: "Plain Text",
   s_loc_recalc: "Recalc",
   s_loc_recalculation: "Recalculation",
   s_loc_redo: "Redo",
   s_loc_right: "Right",
   s_loc_right_border: "Right Border",
   s_loc_sheet_settings: "SHEET SETTINGS",
   s_loc_save: "Save",
   s_loc_save_to: "Save to",
   s_loc_set_cell_contents: "Set Cell Contents",
   s_loc_set_cells_to_sort: "Set Cells To Sort",
   s_loc_set_value_to: "Set Value To",
   s_loc_set_to_link_format: "Set to Link format",
   s_loc_setXclear_move_from: "Set/Clear Move From",
   s_loc_show_cell_settings: "Show Cell Settings",
   s_loc_show_sheet_settings: "Show Sheet Settings",
   s_loc_show_in_new_browser_window: "Show in new browser window",
   s_loc_size: "Size",
   s_loc_socialcalcXsave_format: "SocialCalc-save format",
   s_loc_sort: "Sort",
   s_loc_sort_: "Sort ",
   s_loc_sort_cells: "Sort Cells",
   s_loc_swap_colors: "Swap Colors",
   s_loc_tabXdelimited_format: "Tab-delimited format",
   s_loc_text: "Text",
   s_loc_text_horizontal: "Text Horizontal",
   s_loc_this_is_aXbrXsample: "This is a<br>sample",
   s_loc_top: "Top",
   s_loc_top_border: "Top Border",
   s_loc_undone_steps: "UNDONE STEPS",
   s_loc_url: "URL",
   s_loc_undo: "Undo",
   s_loc_unmerge_cells: "Unmerge Cells",
   s_loc_up: "Up",
   s_loc_value: "Value",
   s_loc_vertical: "Vertical",
   s_loc_wikitext: "Wikitext",
   s_loc_workspace: "Workspace",
   s_loc_XnewX: "[New]",
   s_loc_XnoneX: "[None]",
   s_loc_Xselect_rangeX: "[select range]",

//
// SocialCalc Spreadsheet Viewer module, socialcalcviewer.js:
//

   //*** SocialCalc.SpreadsheetViewer

   SVStatuslineheight: 20, // in pixels
   SVStatuslineCSS: "font-size:10px;padding:3px 0px;",

//
// SocialCalc Format Number module, formatnumber2.js:
//

   FormatNumber_separatorchar: ",", // the thousands separator character when formatting numbers for display
   FormatNumber_decimalchar: ".", // the decimal separator character when formatting numbers for display
   FormatNumber_defaultCurrency: "$", // the currency string used if none specified

   // The following constants are arrays of strings with the short (3 character) and full names of days and months

   s_FormatNumber_daynames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
   s_FormatNumber_daynames3: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
   s_FormatNumber_monthnames: ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                                      "October", "November", "December"],
   s_FormatNumber_monthnames3: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
   s_FormatNumber_am: "AM",
   s_FormatNumber_am1: "A",
   s_FormatNumber_pm: "PM",
   s_FormatNumber_pm1: "P",

//
// SocialCalc Spreadsheet Formula module, formula1.js:
//

   s_parseerrexponent: "Improperly formed number exponent",
   s_parseerrchar: "Unexpected character in formula",
   s_parseerrstring: "Improperly formed string",
   s_parseerrspecialvalue: "Improperly formed special value",
   s_parseerrtwoops: "Error in formula (two operators inappropriately in a row)",
   s_parseerrmissingopenparen: "Missing open parenthesis in list with comma(s). ",
   s_parseerrcloseparennoopen: "Closing parenthesis without open parenthesis. ",
   s_parseerrmissingcloseparen: "Missing close parenthesis. ",
   s_parseerrmissingoperand: "Missing operand. ",
   s_parseerrerrorinformula: "Error in formula.",
   s_calcerrerrorvalueinformula: "Error value in formula",
   s_parseerrerrorinformulabadval: "Error in formula resulting in bad value",
   s_formularangeresult: "Formula results in range value:",
   s_calcerrnumericnan: "Formula results in an bad numeric value",
   s_calcerrnumericoverflow: "Numeric overflow",
   s_sheetunavailable: "Sheet unavailable:", // when FindSheetInCache returns null
   s_calcerrcellrefmissing: "Cell reference missing when expected.",
   s_calcerrsheetnamemissing: "Sheet name missing when expected.",
   s_circularnameref: "Circular name reference to name",
   s_calcerrunknownname: "Unknown name",
   s_calcerrincorrectargstofunction: "Incorrect arguments to function",
   s_sheetfuncunknownfunction: "Unknown function",
   s_sheetfunclnarg: "LN argument must be greater than 0",
   s_sheetfunclog10arg: "LOG10 argument must be greater than 0",
   s_sheetfunclogsecondarg: "LOG second argument must be numeric greater than 0",
   s_sheetfunclogfirstarg: "LOG first argument must be greater than 0",
   s_sheetfuncroundsecondarg: "ROUND second argument must be numeric",
   s_sheetfuncddblife: "DDB life must be greater than 1",
   s_sheetfuncslnlife: "SLN life must be greater than 1",

   // Function definition text

   s_fdef_ABS: 'Absolute value function. ',
   s_fdef_ACOS: 'Trigonometric arccosine function. ',
   s_fdef_AND: 'True if all arguments are true. ',
   s_fdef_ASIN: 'Trigonometric arcsine function. ',
   s_fdef_ATAN: 'Trigonometric arctan function. ',
   s_fdef_ATAN2: 'Trigonometric arc tangent function (result is in radians). ',
   s_fdef_AVERAGE: 'Averages the values. ',
   s_fdef_CHOOSE: 'Returns the value specified by the index. The values may be ranges of cells. ',
   s_fdef_COLUMNS: 'Returns the number of columns in the range. ',
   s_fdef_COS: 'Trigonometric cosine function (value is in radians). ',
   s_fdef_COUNT: 'Counts the number of numeric values, not blank, text, or error. ',
   s_fdef_COUNTA: 'Counts the number of non-blank values. ',
   s_fdef_COUNTBLANK: 'Counts the number of blank values. (Note: "" is not blank.) ',
   s_fdef_COUNTIF: 'Counts the number of number of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). ',
   s_fdef_DATE: 'Returns the appropriate date value given numbers for year, month, and day. For example: DATE(2006,2,1) for February 1, 2006. Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
   s_fdef_DAVERAGE: 'Averages the values in the specified field in records that meet the criteria. ',
   s_fdef_DAY: 'Returns the day of month for a date value. ',
   s_fdef_DCOUNT: 'Counts the number of numeric values, not blank, text, or error, in the specified field in records that meet the criteria. ',
   s_fdef_DCOUNTA: 'Counts the number of non-blank values in the specified field in records that meet the criteria. ',
   s_fdef_DDB: 'Returns the amount of depreciation at the given period of time (the default factor is 2 for double-declining balance).   ',
   s_fdef_DEGREES: 'Converts value in radians into degrees. ',
   s_fdef_DGET: 'Returns the value of the specified field in the single record that meets the criteria. ',
   s_fdef_DMAX: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_DMIN: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_DPRODUCT: 'Returns the result of multiplying the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_DSTDEV: 'Returns the sample standard deviation of the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_DSTDEVP: 'Returns the standard deviation of the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_DSUM: 'Returns the sum of the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_DVAR: 'Returns the sample variance of the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_DVARP: 'Returns the variance of the numeric values in the specified field in records that meet the criteria. ',
   s_fdef_EVEN: 'Rounds the value up in magnitude to the nearest even integer. ',
   s_fdef_EXACT: 'Returns "true" if the values are exactly the same, including case, type, etc. ',
   s_fdef_EXP: 'Returns e raised to the value power. ',
   s_fdef_FACT: 'Returns factorial of the value. ',
   s_fdef_FALSE: 'Returns the logical value "false". ',
   s_fdef_FIND: 'Returns the starting position within string2 of the first occurrence of string1 at or after "start". If start is omitted, 1 is assumed. ',
   s_fdef_FV: 'Returns the future value of repeated payments of money invested at the given rate for the specified number of periods, with optional present value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
   s_fdef_HLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the row offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. ',
   s_fdef_HOUR: 'Returns the hour portion of a time or date/time value. ',
   s_fdef_IF: 'Results in true-value if logical-expression is TRUE or non-zero, otherwise results in false-value. ',
   s_fdef_INDEX: 'Returns a cell or range reference for the specified row and column in the range. If range is 1-dimensional, then only one of rownum or colnum are needed. If range is 2-dimensional and rownum or colnum are zero, a reference to the range of just the specified column or row is returned. You can use the returned reference value in a range, e.g., sum(A1:INDEX(A2:A10,4)). ',
   s_fdef_INT: 'Returns the value rounded down to the nearest integer (towards -infinity). ',
   s_fdef_IRR: 'Returns the interest rate at which the cash flows in the range have a net present value of zero. Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
   s_fdef_ISBLANK: 'Returns "true" if the value is a reference to a blank cell. ',
   s_fdef_ISERR: 'Returns "true" if the value is of type "Error" but not "NA". ',
   s_fdef_ISERROR: 'Returns "true" if the value is of type "Error". ',
   s_fdef_ISLOGICAL: 'Returns "true" if the value is of type "Logical" (true/false). ',
   s_fdef_ISNA: 'Returns "true" if the value is the error type "NA". ',
   s_fdef_ISNONTEXT: 'Returns "true" if the value is not of type "Text". ',
   s_fdef_ISNUMBER: 'Returns "true" if the value is of type "Number" (including logical values). ',
   s_fdef_ISTEXT: 'Returns "true" if the value is of type "Text". ',
   s_fdef_LEFT: 'Returns the specified number of characters from the text value. If count is omitted, 1 is assumed. ',
   s_fdef_LEN: 'Returns the number of characters in the text value. ',
   s_fdef_LN: 'Returns the natural logarithm of the value. ',
   s_fdef_LOG: 'Returns the logarithm of the value using the specified base. ',
   s_fdef_LOG10: 'Returns the base 10 logarithm of the value. ',
   s_fdef_LOWER: 'Returns the text value with all uppercase characters converted to lowercase. ',
   s_fdef_MATCH: 'Look for the matching value for the given value in the range and return position (the first is 1) in that range. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. If rangelookup is -1, act like 1 but the bracket is match>=value. ',
   s_fdef_MAX: 'Returns the maximum of the numeric values. ',
   s_fdef_MID: 'Returns the specified number of characters from the text value starting from the specified position. ',
   s_fdef_MIN: 'Returns the minimum of the numeric values. ',
   s_fdef_MINUTE: 'Returns the minute portion of a time or date/time value. ',
   s_fdef_MOD: 'Returns the remainder of the first value divided by the second. ',
   s_fdef_MONTH: 'Returns the month part of a date value. ',
   s_fdef_N: 'Returns the value if it is a numeric value otherwise an error. ',
   s_fdef_NA: 'Returns the #N/A error value which propagates through most operations. ',
   s_fdef_NOT: 'Returns FALSE if value is true, and TRUE if it is false. ',
   s_fdef_NOW: 'Returns the current date/time. ',
   s_fdef_NPER: 'Returns the number of periods at which payments invested each period at the given rate with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period) has the given present value. ',
   s_fdef_NPV: 'Returns the net present value of cash flows (which may be individual values and/or ranges) at the given rate. The flows are positive if income, negative if paid out, and are assumed at the end of each period. ',
   s_fdef_ODD: 'Rounds the value up in magnitude to the nearest odd integer. ',
   s_fdef_OR: 'True if any argument is true ',
   s_fdef_PI: 'The value 3.1415926... ',
   s_fdef_PMT: 'Returns the amount of each payment that must be invested at the given rate for the specified number of periods to have the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
   s_fdef_POWER: 'Returns the first value raised to the second value power. ',
   s_fdef_PRODUCT: 'Returns the result of multiplying the numeric values. ',
   s_fdef_PROPER: 'Returns the text value with the first letter of each word converted to uppercase and the others to lowercase. ',
   s_fdef_PV: 'Returns the present value of the given number of payments each invested at the given rate, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
   s_fdef_RADIANS: 'Converts value in degrees into radians. ',
   s_fdef_RATE: 'Returns the rate at which the given number of payments each invested at the given rate has the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
   s_fdef_REPLACE: 'Returns text1 with the specified number of characters starting from the specified position replaced by text2. ',
   s_fdef_REPT: 'Returns the text repeated the specified number of times. ',
   s_fdef_RIGHT: 'Returns the specified number of characters from the text value starting from the end. If count is omitted, 1 is assumed. ',
   s_fdef_ROUND: 'Rounds the value to the specified number of decimal places. If precision is negative, then round to powers of 10. The default precision is 0 (round to integer). ',
   s_fdef_ROWS: 'Returns the number of rows in the range. ',
   s_fdef_SECOND: 'Returns the second portion of a time or date/time value (truncated to an integer). ',
   s_fdef_SIN: 'Trigonometric sine function (value is in radians) ',
   s_fdef_SLN: 'Returns the amount of depreciation at each period of time using the straight-line method. ',
   s_fdef_SQRT: 'Square root of the value ',
   s_fdef_STDEV: 'Returns the sample standard deviation of the numeric values. ',
   s_fdef_STDEVP: 'Returns the standard deviation of the numeric values. ',
   s_fdef_SUBSTITUTE: 'Returns text1 with the all occurrences of oldtext replaced by newtext. If "occurrence" is present, then only that occurrence is replaced. ',
   s_fdef_SUM: 'Adds the numeric values. The values to the sum function may be ranges in the form similar to A1:B5. ',
   s_fdef_SUMIF: 'Sums the numeric values of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). If range2 is present, then range1 is tested and the corresponding range2 value is summed. ',
   s_fdef_SYD: 'Depreciation by Sum of Year\'s Digits method. ',
   s_fdef_T: 'Returns the text value or else a null string. ',
   s_fdef_TAN: 'Trigonometric tangent function (value is in radians) ',
   s_fdef_TIME: 'Returns the time value given the specified hour, minute, and second. ',
   s_fdef_TODAY: 'Returns the current date (an integer). Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
   s_fdef_TRIM: 'Returns the text value with leading, trailing, and repeated spaces removed. ',
   s_fdef_TRUE: 'Returns the logical value "true". ',
   s_fdef_TRUNC: 'Truncates the value to the specified number of decimal places. If precision is negative, truncate to powers of 10. ',
   s_fdef_UPPER: 'Returns the text value with all lowercase characters converted to uppercase. ',
   s_fdef_VALUE: 'Converts the specified text value into a numeric value. Various forms that look like numbers (including digits followed by %, forms that look like dates, etc.) are handled. This may not handle all of the forms accepted by other spreadsheets and may be locale dependent. ',
   s_fdef_VAR: 'Returns the sample variance of the numeric values. ',
   s_fdef_VARP: 'Returns the variance of the numeric values. ',
   s_fdef_VLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the column offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match>=value) instead of exact match. ',
   s_fdef_WEEKDAY: 'Returns the day of week specified by the date value. If type is 1 (the default), Sunday is day and Saturday is day 7. If type is 2, Monday is day 1 and Sunday is day 7. If type is 3, Monday is day 0 and Sunday is day 6. ',
   s_fdef_YEAR: 'Returns the year part of a date value. ',
   s_fdef_SUMPRODUCT: 'Sums the pairwise products of 2 or more ranges. The ranges must be of equal length.',
   s_fdef_CEILING: 'Rounds the given number up to the nearest integer or multiple of significance. Significance is the value to whose multiple of ten the value is to be rounded up (.01, .1, 1, 10, etc.)',
   s_fdef_FLOOR: 'Rounds the given number down to the nearest multiple of significance. Significance is the value to whose multiple of ten the number is to be rounded down (.01, .1, 1, 10, etc.)',

   s_farg_v: "value",
   s_farg_vn: "value1, value2, ...",
   s_farg_xy: "valueX, valueY",
   s_farg_choose: "index, value1, value2, ...",
   s_farg_range: "range",
   s_farg_rangec: "range, criteria",
   s_farg_date: "year, month, day",
   s_farg_dfunc: "databaserange, fieldname, criteriarange",
   s_farg_ddb: "cost, salvage, lifetime, period, [factor]",
   s_farg_find: "string1, string2, [start]",
   s_farg_fv: "rate, n, payment, [pv, [paytype]]",
   s_farg_hlookup: "value, range, row, [rangelookup]",
   s_farg_iffunc: "logical-expression, true-value, [false-value]",
   s_farg_index: "range, rownum, colnum",
   s_farg_irr: "range, [guess]",
   s_farg_tc: "text, count",
   s_farg_log: "value, base",
   s_farg_match: "value, range, [rangelookup]",
   s_farg_mid: "text, start, length",
   s_farg_nper: "rate, payment, pv, [fv, [paytype]]",
   s_farg_npv: "rate, value1, value2, ...",
   s_farg_pmt: "rate, n, pv, [fv, [paytype]]",
   s_farg_pv: "rate, n, payment, [fv, [paytype]]",
   s_farg_rate: "n, payment, pv, [fv, [paytype, [guess]]]",
   s_farg_replace: "text1, start, length, text2",
   s_farg_vp: "value, [precision]",
   s_farg_valpre: "value, precision",
   s_farg_csl: "cost, salvage, lifetime",
   s_farg_cslp: "cost, salvage, lifetime, period",
   s_farg_subs: "text1, oldtext, newtext, [occurrence]",
   s_farg_sumif: "range1, criteria, [range2]",
   s_farg_hms: "hour, minute, second",
   s_farg_txt: "text",
   s_farg_vlookup: "value, range, col, [rangelookup]",
   s_farg_weekday: "date, [type]",
   s_farg_dt: "date",
   s_farg_rangen: "range1, range2, ...",
   s_farg_vsig: 'value, [significance]',

   function_classlist: ["all", "stat", "lookup", "datetime", "financial", "test", "math", "text"], // order of function classes

   s_fclass_all: "All",
   s_fclass_stat: "Statistics",
   s_fclass_lookup: "Lookup",
   s_fclass_datetime: "Date & Time",
   s_fclass_financial: "Financial",
   s_fclass_test: "Test",
   s_fclass_math: "Math",
   s_fclass_text: "Text",

   lastone: null

   };

// Default classnames for use with SocialCalc.ConstantsSetClasses:

SocialCalc.ConstantsDefaultClasses = {
   defaultComment: "",
   defaultCommentNoGrid: "",
   defaultHighlightTypeCursor: "",
   defaultHighlightTypeRange: "",
   defaultColname: "",
   defaultSelectedColname: "",
   defaultRowname: "",
   defaultSelectedRowname: "", 
   defaultUpperLeft: "",
   defaultSkippedCell: "",
   defaultPaneDivider: "",
   cteGriddiv: "", // this one has no Style version with it
   defaultInputEcho: {classname: "", style: "filter:alpha(opacity=90);opacity:.9;"}, // so FireFox won't show warning
   TCmain: "",
   TCendcap: "",
   TCpaneslider: "",
   TClessbutton: "",
   TCmorebutton: "",
   TCscrollarea: "",
   TCthumb: "",
   TCPStrackingline: "",
   TCTDFSthumbstatus: "",
   TDpopupElement: ""
   };

//
// SocialCalc.ConstantsSetClasses(prefix)
//
// This routine goes through all of the xyzClass/xyzStyle pairs and sets the Class to a default and
// turns off the Style, if present. The prefix is put before each default.
// The list of items to set is in SocialCalc.ConstantsDefaultClasses. The names there
// correspond to the "xyz" parts. If there is a value, it is the default to set. If the
// default is a null, no change is made. If the default is the null string (""), the
// name of the item is used (e.g., "defaultComment" would use the classname "defaultComment").
// If the default is an object, then it expects {classname: classname, style: stylestring} - this
// lets you combine both.
//

SocialCalc.ConstantsSetClasses = function(prefix) {

   var defaults = SocialCalc.ConstantsDefaultClasses;
   var scc = SocialCalc.Constants;
   var item;

   prefix = prefix || "";

   for (item in defaults) {
      if (typeof defaults[item] == "string") {
         scc[item+"Class"] = prefix + (defaults[item] || item);
         if (scc[item+"Style"] !== undefined) {
            scc[item+"Style"] = "";
            }
         }
      else if (typeof defaults[item] == "object") {
         scc[item+"Class"] = prefix + (defaults[item].classname || item);
         scc[item+"Style"] = defaults[item].style;
         }
      }
   }

// Set the image prefix on all images.

SocialCalc.ConstantsSetImagePrefix = function(imagePrefix) {

   var scc = SocialCalc.Constants;

   for (var item in scc) {
      if (typeof scc[item] == "string") {
         scc[item] = scc[item].replace(scc.defaultImagePrefix, imagePrefix);
         }
      }
   scc.defaultImagePrefix = imagePrefix;

   }

//
// The main SocialCalc code module of the SocialCalc package
//
/*
// (c) Copyright 2010 Socialtext, Inc.
// All Rights Reserved.
//
// The contents of this file are subject to the Artistic License 2.0; you may not
// use this file except in compliance with the License. You may obtain a copy of 
// the License at http://socialcalc.org/licenses/al-20/.
//
// Some of the other files in the SocialCalc package are licensed under
// different licenses. Please note the licenses of the modules you use.
//
// Code History:
//
// Initially coded by Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
// Based in part on the SocialCalc 1.1.0 code written in Perl.
// The SocialCalc 1.1.0 code was:
//    Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc.
//    All Rights Reserved.
//    Portions (c) Copyright 2007 Socialtext, Inc.
//    All Rights Reserved.
// The Perl SocialCalc started as modifications to the wikiCalc(R) program, version 1.0.
// wikiCalc 1.0 was written by Software Garden, Inc.
// Unless otherwise specified, referring to "SocialCalc" in comments refers to this
// JavaScript version of the code, not the SocialCalc Perl code.
//
*/

/*

**** Overview ****

This is the beginning of a library of routines for displaying and editing spreadsheet
data in a browser. The HTML that includes this does not need to have anything
specific to the spreadsheet or editor already present -- everything is dynamically
added to the DOM by this code, including the rendered sheet and any editing controls.

The library has a few parts. This is the main SocialCalc code module.
Other parts are the Table Editor module, the Formula module, and the Format Number module.
Note: The Table Editor module is licensed under a different license than this module.

The class/object style is derived from O'Reilly's JavaScript by Flanagan, 5th Edition,
section 9.3, page 157.

All of the data, object definitions, functions, etc., are stored as properties of the SocialCalc
object so as not to clutter up the global variables nor conflict with other names.

A design goal (not tested yet for success) is to make it possible to have more than one
spreadsheet active on a page, perhaps even open for editing. It is assumed, though, that
there is only one mouse and one keyboard (a good assumption on most PCs today but not in the
new "touch and surface world" Apple and Microsoft are working towards).

The testing has been on Windows Firefox (2 and 3),
Internet Explorer (6 and 7), Opera (9.23 and mainly later), Mac Safari (3.1), and Mac Firefox (2.0.0.6).
There are small issues with Firefox before 2.0 (cosmetic with drag handles) and larger ones
with Opera before 9.5 (the Delete key isn't recognized in some cases -- the 9.5 version was still
in beta and this bug affects other products like GMail, I believe).

The data is stored in a SocialCalc.Sheet object. The data is organized in a form similar to that
used by SocialCalc 1.1.0. There is a function for converting a normal SocialCalc spreadsheet
save data string (the spreadsheet part of a SocialCalc data file) into this internal form.

The SocialCalc.RenderContext class provides methods for rendering a table into the DOM representing
part of the spreadsheet. It is assumed that the spreadsheet could possibly be very large
and that rendering the whole thing at once could be too time consuming. It is also set up so
that it might be possible to have some of the sheet data only be loaded on demand (such as by Ajax).
The rendering can render cells to the right and below the already active area of the spreadsheet
so that you can scroll to that "clean" area without explicitly doing "add row/column". The class also
does simple operations such as "scrolling" within that table. The table may optionally include
row and column headers and may be split into panes. Most of the code assumes any number of panes,
but only the rightmost pane has scrolling code. In normal operation there would be one or two
panes horizontally and vertically. The panes may start on any row/column, though a given row/column
should only appear in one pane at a time (not all code enforces this, yet).

The RenderContext is designed to be rendered as part of a SocialCalc.TableEditor. The TableEditor
includes the spreadsheet grid as well as scrollbars, pane sliders, and (eventually) editing controls.
The layout is dynamic and may be recomputed on the fly, such as in response to resizing the browser
window.

The scrollbars and pane sliders are created using SocialCalc.TableControl objects. These in turn
make use of Dragging, ToolTip, Button, and MouseWheel functions.

The keyboard input is handled by keyboard code.

There are also some helper routines.

More comments yet to come...

*/


var SocialCalc;
if (!SocialCalc) SocialCalc = {};

// *************************************
//
// Shared values
//
// These are "global" values shared by the classes, including default settings
//
// *************************************

// Callbacks

SocialCalc.Callbacks = {

   // The next two are used by SocialCalc.format_text_for_display

   // The function to expand wiki text - should be set if you want wikitext expansion
   // The form is: expand_wiki(displayvalue, sheetobj, linkstyle, valueformat)
   //    valueformat is text-wiki followed by optional sub-formats, e.g., text-wikipagelink

   expand_wiki: null,

   expand_markup: function(displayvalue, sheetobj, linkstyle) // the old function to expand wiki text - may be replaced
                   {return SocialCalc.default_expand_markup(displayvalue, sheetobj, linkstyle);},

   // MakePageLink is used to create the href for a link to another "page"
   // The form is: MakePageLink(pagename, workspacename, linktyle, valueformat), returns string

   MakePageLink: null,

   // NormalizeSheetName is used to make different variations of sheetnames use the same cache slot

   NormalizeSheetName: null // use default - lowercase

   };

// Shared flags

   // none at present


// *************************************
//
// Cell class:
//
// *************************************

//
// Class SocialCalc.Cell
//
// Usage: var s = new SocialCalc.Cell(coord);
//
// Cell attributes include:
//
//    coord: the column/row as a string, e.g., "A1"
//    datavalue: the value to be used for computation and formatting for display,
//               string or numeric (tolerant of numbers stored as strings)
//    datatype: if present, v=numeric value, t=text value, f=formula,
//              or c=constant that is not a simple number (like "$1.20")
//    formula: if present, the formula (without leading "=") for computation or the constant
//    valuetype: first char is main type, the following are sub-types.
//               Main types are b=blank cell, n=numeric, t=text, e=error
//               Examples of using sub-types would be "nt" for a numeric time value, "n$" for currency, "nl" for logical
//    readonly: if present, whether the current cell is read-only of writable
//    displayvalue: if present, rendered version of datavalue with formatting attributes applied
//    parseinfo: if present, cached parsed version of formula
//
//    The following optional values, if present, are mainly used in rendering, overriding defaults:
//
//    bt, br, bb, bl: number of border's definition
//    layout: layout (vertical alignment, padding) definition number
//    font: font definition number
//    color: text color definition number
//    bgcolor: background color definition number
//    cellformat: cell format (horizontal alignment) definition number
//    nontextvalueformat: custom format definition number for non-text values, e.g., numbers
//    textvalueformat: custom format definition number for text values
//    colspan, rowspan: number of cells to span for merged cells (only on main cell)
//    cssc: custom css classname for cell, as text (no special chars)
//    csss: custom css style definition
//    mod: modification allowed flag "y" if present
//    comment: cell comment string
//

SocialCalc.Cell = function(coord) {

   this.coord = coord;
   this.datavalue = "";
   this.datatype = null;
   this.formula = "";
   this.valuetype = "b";
   this.readonly = false;

   }

// The types of cell properties
//
// Type 1: Base, Type 2: Attribute, Type 3: Special (e.g., displaystring, parseinfo)

SocialCalc.CellProperties = {
   coord: 1, datavalue: 1, datatype: 1, formula: 1, valuetype: 1, errors: 1, comment: 1, readonly: 1,
   bt: 2, br: 2, bb: 2, bl: 2, layout: 2, font: 2, color: 2, bgcolor: 2,
   cellformat: 2, nontextvalueformat: 2, textvalueformat: 2, colspan: 2, rowspan: 2,
   cssc: 2, csss: 2, mod: 2,
   displaystring: 3, // used to cache rendered HTML of cell contents
   parseinfo: 3, // used to cache parsed formulas
   hcolspan: 3, hrowspan: 3 // spans taking hidden cols/rows into account (!!! NOT YET !!!)
   };

SocialCalc.CellPropertiesTable = {
   bt: "borderstyle", br: "borderstyle", bb: "borderstyle", bl: "borderstyle",
   layout: "layout", font: "font", color: "color", bgcolor: "color",
   cellformat: "cellformat", nontextvalueformat: "valueformat", textvalueformat: "valueformat"
   };

// *************************************
//
// Sheet class:
//
// *************************************

//
// Class SocialCalc.Sheet
//
// Usage: var s = new SocialCalc.Sheet();
//

SocialCalc.Sheet = function() {

   SocialCalc.ResetSheet(this);

   // Set other values:
   //
   // sheet.statuscallback(data, status, arg, this.statuscallbackparams) is called
   // during recalc and commands.
   //
   // During recalc, data is the current recalcdata.
   // The values for status and the corresponding arg are:
   //
   //    calcorder, {coord: coord, total: celllist length, count: count} [0 or more times per recalc]
   //    calccheckdone, calclist length [once per recalc]
   //    calcstep, {coord: coord, total: calclist length, count: count} [0 or more times per recalc]
   //    calcloading, {sheetname: name-of-sheet}
   //    calcserverfunc, {funcname: name-of-function, coord: coord, total: calclist length, count: count}
   //    calcfinished, time in milliseconds [once per recalc]
   //
   // During commands, data is SocialCalc.SheetCommandInfo.
   // These values for status and arg are:
   //
   //    cmdstart, cmdstr
   //    cmdend
   //

   this.statuscallback = null; // routine called with cmdstart, calcstart, etc., status and args:
                                // sheet.statuscallback(data, status, arg, params)
   this.statuscallbackparams = null; // parameters passed to that routine

   }

//
// SocialCalc.ResetSheet(sheet)
//
// Resets (and/or initializes) sheet data values.
// 

SocialCalc.ResetSheet = function(sheet, reload) {

   // properties:

   sheet.cells = {}; // at least one for each non-blank cell: coord: cell-object
   sheet.attribs = // sheet attributes
      {
         lastcol: 1,
         lastrow: 1,
         defaultlayout: 0,
         usermaxcol: 0,
         usermaxrow: 0

      };
   sheet.rowattribs =
      {
         hide: {}, // access by row number
         height: {}
      };
   sheet.colattribs =
      {
         width: {}, // access by col name
         hide: {}
      };
   sheet.names={}; // Each is: {desc: "optional description", definition: "B5, A1:B7, or =formula"}
   sheet.layouts=[];
   sheet.layouthash={};
   sheet.fonts=[];
   sheet.fonthash={};
   sheet.colors=[];
   sheet.colorhash={};
   sheet.borderstyles=[];
   sheet.borderstylehash={};
   sheet.cellformats=[];
   sheet.cellformathash={};
   sheet.valueformats=[];
   sheet.valueformathash={};

   sheet.copiedfrom = ""; // if a range, then this was loaded from a saved range as clipboard content

   sheet.changes = new SocialCalc.UndoStack();

   sheet.renderneeded = false;

   sheet.changedrendervalues = true; // if true, spans and/or fonts have changed (set by ExecuteSheetCommand & GetStyle)

   sheet.recalcchangedavalue = false; // true if a recalc resulted in a change to a cell's calculated value

   sheet.hiddencolrow = ""; // "col" or "row" if it was hidden

   sheet.sci = new SocialCalc.SheetCommandInfo(sheet);

   }

// Methods:

SocialCalc.Sheet.prototype.ResetSheet = function() {SocialCalc.ResetSheet(this);};
SocialCalc.Sheet.prototype.AddCell = function(newcell) {return this.cells[newcell.coord]=newcell;};
SocialCalc.Sheet.prototype.GetAssuredCell = function(coord) {
   return this.cells[coord] || this.AddCell(new SocialCalc.Cell(coord));
   };
SocialCalc.Sheet.prototype.ParseSheetSave = function(savedsheet) {SocialCalc.ParseSheetSave(savedsheet,this);};
SocialCalc.Sheet.prototype.CellFromStringParts = function(cell, parts, j) {return SocialCalc.CellFromStringParts(this, cell, parts, j);};
SocialCalc.Sheet.prototype.CreateSheetSave = function(range, canonicalize) {return SocialCalc.CreateSheetSave(this, range, canonicalize);};
SocialCalc.Sheet.prototype.CellToString = function(cell) {return SocialCalc.CellToString(this, cell);};
SocialCalc.Sheet.prototype.CanonicalizeSheet = function(full) {return SocialCalc.CanonicalizeSheet(this, full);};
SocialCalc.Sheet.prototype.EncodeCellAttributes = function(coord) {return SocialCalc.EncodeCellAttributes(this, coord);};
SocialCalc.Sheet.prototype.EncodeSheetAttributes = function() {return SocialCalc.EncodeSheetAttributes(this);};
SocialCalc.Sheet.prototype.DecodeCellAttributes = function(coord, attribs, range) {return SocialCalc.DecodeCellAttributes(this, coord, attribs, range);};
SocialCalc.Sheet.prototype.DecodeSheetAttributes = function(attribs) {return SocialCalc.DecodeSheetAttributes(this, attribs);};

SocialCalc.Sheet.prototype.ScheduleSheetCommands = function(cmd, saveundo) {return SocialCalc.ScheduleSheetCommands(this, cmd, saveundo);};
SocialCalc.Sheet.prototype.SheetUndo = function() {return SocialCalc.SheetUndo(this);};
SocialCalc.Sheet.prototype.SheetRedo = function() {return SocialCalc.SheetRedo(this);};
SocialCalc.Sheet.prototype.CreateAuditString = function() {return SocialCalc.CreateAuditString(this);};
SocialCalc.Sheet.prototype.GetStyleNum = function(atype, style) {return SocialCalc.GetStyleNum(this, atype, style);};
SocialCalc.Sheet.prototype.GetStyleString = function(atype, num) {return SocialCalc.GetStyleString(this, atype, num);};
SocialCalc.Sheet.prototype.RecalcSheet = function() {return SocialCalc.RecalcSheet(this);};

//
// Sheet save format:
//
// linetype:param1:param2:...
//
// Linetypes are:
//
//    version:versionname - version of this format. Currently 1.4.
//
//    cell:coord:type:value...:type:value... - Types are as follows:
//
//       v:value - straight numeric value
//       t:value - straight text/wiki-text in cell, encoded to handle \, :, newlines
//       vt:fulltype:value - value with value type/subtype
//       vtf:fulltype:value:formulatext - formula resulting in value with value type/subtype, value and text encoded
//       vtc:fulltype:value:valuetext - formatted text constant resulting in value with value type/subtype, value and text encoded
//       vf:fvalue:formulatext - formula resulting in value, value and text encoded (obsolete: only pre format version 1.1)
//          fvalue - first char is "N" for numeric value, "T" for text value, "H" for HTML value, rest is the value
//       e:errortext - Error text. Non-blank means formula parsing/calculation results in error.
//       b:topborder#:rightborder#:bottomborder#:leftborder# - border# in sheet border list or blank if none
//       l:layout# - number in cell layout list
//       f:font# - number in sheet fonts list
//       c:color# - sheet color list index for text
//       bg:color# - sheet color list index for background color
//       cf:format# - sheet cell format number for explicit format (align:left, etc.)
//       cvf:valueformat# - sheet cell value format number (obsolete: only pre format v1.2)
//       tvf:valueformat# - sheet cell text value format number
//       ntvf:valueformat# - sheet cell non-text value format number
//       colspan:numcols - number of columns spanned in merged cell
//       rowspan:numrows - number of rows spanned in merged cell
//       cssc:classname - name of CSS class to be used for cell when published instead of one calculated here
//       csss:styletext - explicit CSS style information, encoded to handle :, etc.
//       mod:allow - if "y" allow modification of cell for live "view" recalc
//       comment:value - encoded text of comment for this cell (added in v1.5)
//
//    col:
//       w:widthval - number, "auto" (no width in <col> tag), number%, or blank (use default)
//       hide: - yes/no, no is assumed if missing
//    row:
//       hide - yes/no, no is assumed if missing
//
//    sheet:
//       c:lastcol - number
//       r:lastrow - number
//       w:defaultcolwidth - number, "auto", number%, or blank (default->80)
//       h:defaultrowheight - not used
//       tf:format# - cell format number for sheet default for text values
//       ntf:format# - cell format number for sheet default for non-text values (i.e., numbers)
//       layout:layout# - default cell layout number in cell layout list
//       font:font# - default font number in sheet font list
//       vf:valueformat# - default number value format number in sheet valueformat list (obsolete: only pre format version 1.2)
//       ntvf:valueformat# - default non-text (number) value format number in sheet valueformat list
//       tvf:valueformat# - default text value format number in sheet valueformat list
//       color:color# - default number for text color in sheet color list
//       bgcolor:color# - default number for background color in sheet color list
//       circularreferencecell:coord - cell coord with a circular reference
//       recalc:value - on/off (on is default). If not "off", appropriate changes to the sheet cause a recalc
//       needsrecalc:value - yes/no (no is default). If "yes", formula values are not up to date
//       usermaxcol:value - maximum column to display, 0 for unlimited (default=0)
//       usermaxrow:value - maximum row to display, 0 for unlimited (default=0)
//
//    name:name:description:value - name definition, name in uppercase, with value being "B5", "A1:B7", or "=formula";
//                                  description and value are encoded.
//    font:fontnum:value - text of font definition (style weight size family) for font fontnum
//                         "*" for "style weight", size, or family, means use default (first look to sheet, then builtin)
//    color:colornum:rgbvalue - text of color definition (e.g., rgb(255,255,255)) for color colornum
//    border:bordernum:value - text of border definition (thickness style color) for border bordernum
//    layout:layoutnum:value - text of vertical alignment and padding style for cell layout layoutnum (* for default):
//                             vertical-alignment:vavalue;padding:topval rightval bottomval leftval;
//    cellformat:cformatnum:value - text of cell alignment (left/center/right) for cellformat cformatnum
//    valueformat:vformatnum:value - text of number format (see FormatValueForDisplay) for valueformat vformatnum (changed in v1.2)
//    clipboardrange:upperleftcoord:bottomrightcoord - ignored -- from wikiCalc
//    clipboard:coord:type:value:... - ignored -- from wikiCalc
//
// If this is clipboard contents, then there is also information to facilitate pasting:
//
//    copiedfrom:upperleftcoord:bottomrightcoord - range from which this was copied
//

// Functions:

SocialCalc.ParseSheetSave = function(savedsheet,sheetobj) {

   var lines=savedsheet.split(/\r\n|\n/);
   var parts=[];
   var line;
   var i, j, t, v, coord, cell, attribs, name;
   var scc = SocialCalc.Constants;

   for (i=0;i<lines.length;i++) {
      line=lines[i];
      parts = line.split(":");
      switch (parts[0]) {
         case "cell":
            cell=sheetobj.GetAssuredCell(parts[1]);
            j=2;
            sheetobj.CellFromStringParts(cell, parts, j);
            break;

         case "col":
            coord=parts[1];
            j=2;
            while (t=parts[j++]) {
               switch (t) {
                  case "w":
                     sheetobj.colattribs.width[coord]=parts[j++]; // must be text - could be auto or %, etc.
                     break;
                  case "hide":
                     sheetobj.colattribs.hide[coord]=parts[j++];
                     break;
                  default:
                     throw scc.s_pssUnknownColType+" '"+t+"'";
                     break;
                  }
               }
            break;

         case "row":
            coord=parts[1]-0;
            j=2;
            while (t=parts[j++]) {
               switch (t) {
                  case "h":
                     sheetobj.rowattribs.height[coord]=parts[j++]-0;
                     break;
                  case "hide":
                     sheetobj.rowattribs.hide[coord]=parts[j++];
                     break;
                  default:
                     throw scc.s_pssUnknownRowType+" '"+t+"'";
                     break;
                  }
               }
            break;

         case "sheet":
            attribs=sheetobj.attribs;
            j=1;
            while (t=parts[j++]) {
               switch (t) {
                  case "c":
                     attribs.lastcol=parts[j++]-0;
                     break;
                  case "r":
                     attribs.lastrow=parts[j++]-0;
                     break;
                  case "w":
                     attribs.defaultcolwidth=parts[j++]+"";
                     break;
                  case "h":
                     attribs.defaultrowheight=parts[j++]-0;
                     break;
                  case "tf":
                     attribs.defaulttextformat=parts[j++]-0;
                     break;
                  case "ntf":
                     attribs.defaultnontextformat=parts[j++]-0;
                     break;
                  case "layout":
                     attribs.defaultlayout=parts[j++]-0;
                     break;
                  case "font":
                     attribs.defaultfont=parts[j++]-0;
                     break;
                  case "tvf":
                     attribs.defaulttextvalueformat=parts[j++]-0;
                     break;
                  case "ntvf":
                     attribs.defaultnontextvalueformat=parts[j++]-0;
                     break;
                  case "color":
                     attribs.defaultcolor=parts[j++]-0;
                     break;
                  case "bgcolor":
                     attribs.defaultbgcolor=parts[j++]-0;
                     break;
                  case "circularreferencecell":
                     attribs.circularreferencecell=parts[j++];
                     break;
                  case "recalc":
                     attribs.recalc=parts[j++];
                     break;
                  case "needsrecalc":
                     attribs.needsrecalc=parts[j++];
                     break;
                  case "usermaxcol":
                     attribs.usermaxcol=parts[j++]-0;
                     break;
                  case "usermaxrow":
                     attribs.usermaxrow=parts[j++]-0;
                     break;
                  default:
                     j+=1;
                     break;
                  }
               }
            break;

         case "name":
            name = SocialCalc.decodeFromSave(parts[1]).toUpperCase();
            sheetobj.names[name] = {desc: SocialCalc.decodeFromSave(parts[2])};
            sheetobj.names[name].definition = SocialCalc.decodeFromSave(parts[3]);
            break;

         case "layout":
            parts=lines[i].match(/^layout\:(\d+)\:(.+)$/); // layouts can have ":" in them
            sheetobj.layouts[parts[1]-0]=parts[2];
            sheetobj.layouthash[parts[2]]=parts[1]-0;
            break;

         case "font":
            sheetobj.fonts[parts[1]-0]=parts[2];
            sheetobj.fonthash[parts[2]]=parts[1]-0;
            break;

         case "color":
            sheetobj.colors[parts[1]-0]=parts[2];
            sheetobj.colorhash[parts[2]]=parts[1]-0;
            break;

         case "border":
            sheetobj.borderstyles[parts[1]-0]=parts[2];
            sheetobj.borderstylehash[parts[2]]=parts[1]-0;
            break;

         case "cellformat":
            v=SocialCalc.decodeFromSave(parts[2]);
            sheetobj.cellformats[parts[1]-0]=v;
            sheetobj.cellformathash[v]=parts[1]-0;
            break;

         case "valueformat":
            v=SocialCalc.decodeFromSave(parts[2]);
            sheetobj.valueformats[parts[1]-0]=v;
            sheetobj.valueformathash[v]=parts[1]-0;
            break;

         case "version":
            break;

         case "copiedfrom":
            sheetobj.copiedfrom = parts[1]+":"+parts[2];
            break;

         case "clipboardrange": // in save versions up to 1.3. Ignored.
         case "clipboard":
            break;

         case "":
            break;

         default:
alert(scc.s_pssUnknownLineType+" '"+parts[0]+"'");
            throw scc.s_pssUnknownLineType+" '"+parts[0]+"'";
            break;
         }
      parts = null;
      }

   }

//
// SocialCalc.CellFromStringParts(sheet, cell, parts, j)
//
// Takes string that has been split by ":" in parts, starting at item j,
// and fills in cell assuming save format.
//

SocialCalc.CellFromStringParts = function(sheet, cell, parts, j) {

   var cell, t, v;

   while (t=parts[j++]) {
      switch (t) {
         case "v":
            cell.datavalue=SocialCalc.decodeFromSave(parts[j++])-0;
            cell.datatype="v";
            cell.valuetype="n";
            break;
         case "t":
            cell.datavalue=SocialCalc.decodeFromSave(parts[j++]);
            cell.datatype="t";
            cell.valuetype=SocialCalc.Constants.textdatadefaulttype; 
            break;
         case "vt":
            v=parts[j++];
            cell.valuetype=v;
            if (v.charAt(0)=="n") {
               cell.datatype="v";
               cell.datavalue=SocialCalc.decodeFromSave(parts[j++])-0;
               }
            else {
               cell.datatype="t";
               cell.datavalue=SocialCalc.decodeFromSave(parts[j++]);
               }
            break;
         case "vtf":
            v=parts[j++];
            cell.valuetype=v;
            if (v.charAt(0)=="n") {
               cell.datavalue=SocialCalc.decodeFromSave(parts[j++])-0;
               }
            else {
               cell.datavalue=SocialCalc.decodeFromSave(parts[j++]);
               }
            cell.formula=SocialCalc.decodeFromSave(parts[j++]);
            cell.datatype="f";
            break;
         case "vtc":
            v=parts[j++];
            cell.valuetype=v;
            if (v.charAt(0)=="n") {
               cell.datavalue=SocialCalc.decodeFromSave(parts[j++])-0;
               }
            else {
               cell.datavalue=SocialCalc.decodeFromSave(parts[j++]);
               }
            cell.formula=SocialCalc.decodeFromSave(parts[j++]);
            cell.datatype="c";
            break;
         case "ro":
            ro=SocialCalc.decodeFromSave(parts[j++]);
            cell.readonly=ro.toLowerCase()=="yes";
            break;
         case "e":
            cell.errors=SocialCalc.decodeFromSave(parts[j++]);
            break;
         case "b":
            cell.bt=parts[j++]-0;
            cell.br=parts[j++]-0;
            cell.bb=parts[j++]-0;
            cell.bl=parts[j++]-0;
            break;
         case "l":
            cell.layout=parts[j++]-0;
            break;
         case "f":
            cell.font=parts[j++]-0;
            break;
         case "c":
            cell.color=parts[j++]-0;
            break;
         case "bg":
            cell.bgcolor=parts[j++]-0;
            break;
         case "cf":
            cell.cellformat=parts[j++]-0;
            break;
         case "ntvf":
            cell.nontextvalueformat=parts[j++]-0;
            break;
         case "tvf":
            cell.textvalueformat=parts[j++]-0;
            break;
         case "colspan":
            cell.colspan=parts[j++]-0;
            break;
         case "rowspan":
            cell.rowspan=parts[j++]-0;
            break;
         case "cssc":
            cell.cssc=parts[j++];
            break;
         case "csss":
            cell.csss=SocialCalc.decodeFromSave(parts[j++]);
            break;
         case "mod":
            j+=1;
            break;
         case "comment":
            cell.comment=SocialCalc.decodeFromSave(parts[j++]);
            break;
         default:
            throw SocialCalc.Constants.s_cfspUnknownCellType+" '"+t+"'";
            break;
         }
      }

   }


SocialCalc.sheetfields = ["defaultrowheight", "defaultcolwidth", "circularreferencecell", "recalc", "needsrecalc", "usermaxcol", "usermaxrow"];
SocialCalc.sheetfieldsshort = ["h", "w", "circularreferencecell", "recalc", "needsrecalc", "usermaxcol", "usermaxrow"];

SocialCalc.sheetfieldsxlat = ["defaulttextformat", "defaultnontextformat",
                              "defaulttextvalueformat", "defaultnontextvalueformat",
                              "defaultcolor", "defaultbgcolor", "defaultfont", "defaultlayout"];
SocialCalc.sheetfieldsxlatshort = ["tf", "ntf", "tvf", "ntvf", "color", "bgcolor", "font", "layout"];
SocialCalc.sheetfieldsxlatxlt = ["cellformat", "cellformat", "valueformat", "valueformat",
                                  "color", "color", "font", "layout"];

//
// sheetstr = SocialCalc.CreateSheetSave(sheetobj, range, canonicalize)
//
// Creates a text representation of the sheetobj data.
// If the range is present then only those cells are saved
// (as clipboard data with "copiedfrom" set).
//

SocialCalc.CreateSheetSave = function(sheetobj, range, canonicalize) {

   var cell, cr1, cr2, row, col, coord, attrib, line, value, formula, i, t, r, b, l, name, blanklen;
   var result=[];

   var prange;

   sheetobj.CanonicalizeSheet(canonicalize || SocialCalc.Constants.doCanonicalizeSheet);
   var xlt = sheetobj.xlt;

   if (range) {
      prange = SocialCalc.ParseRange(range);
      }
   else {
      prange = {cr1: {row: 1, col:1},
                cr2: {row: xlt.maxrow, col: xlt.maxcol}};
      }
   cr1 = prange.cr1;
   cr2 = prange.cr2;

   result.push("version:1.5");

   for (row=cr1.row; row <= cr2.row; row++) {
      for (col=cr1.col; col <= cr2.col; col++) {
         coord = SocialCalc.crToCoord(col, row);
         cell=sheetobj.cells[coord];
         if (!cell) continue;
         line=sheetobj.CellToString(cell);
         if (line.length==0) continue; // ignore completely empty cells
         line="cell:"+coord+line;
         result.push(line);
         }
      }

   for (col=1; col <= xlt.maxcol; col++) {
      coord = SocialCalc.rcColname(col);
      if (sheetobj.colattribs.width[coord])
         result.push("col:"+coord+":w:"+sheetobj.colattribs.width[coord]);
      if (sheetobj.colattribs.hide[coord])
         result.push("col:"+coord+":hide:"+sheetobj.colattribs.hide[coord]);
      }

   for (row=1; row <= xlt.maxrow; row++) {
      if (sheetobj.rowattribs.height[row])
         result.push("row:"+row+":h:"+sheetobj.rowattribs.height[row]);
      if (sheetobj.rowattribs.hide[row])
         result.push("row:"+row+":hide:"+sheetobj.rowattribs.hide[row]);
      }

   line="sheet:c:"+xlt.maxcol+":r:"+xlt.maxrow;

   for (i=0; i<SocialCalc.sheetfields.length; i++) { // non-xlated values
      value = SocialCalc.encodeForSave(sheetobj.attribs[SocialCalc.sheetfields[i]]);
      if (value) line+=":"+SocialCalc.sheetfieldsshort[i]+":"+value;
      }
   for (i=0; i<SocialCalc.sheetfieldsxlat.length; i++) { // xlated values
      value = sheetobj.attribs[SocialCalc.sheetfieldsxlat[i]];
      if (value) line+=":"+SocialCalc.sheetfieldsxlatshort[i]+":"+xlt[SocialCalc.sheetfieldsxlatxlt[i]+"sxlat"][value];
      }

   result.push(line);

   for (i=1;i<xlt.newborderstyles.length;i++) {
      result.push("border:"+i+":"+xlt.newborderstyles[i]);
      }

   for (i=1;i<xlt.newcellformats.length;i++) {
      result.push("cellformat:"+i+":"+SocialCalc.encodeForSave(xlt.newcellformats[i]));
      }

   for (i=1;i<xlt.newcolors.length;i++) {
      result.push("color:"+i+":"+xlt.newcolors[i]);
      }

   for (i=1;i<xlt.newfonts.length;i++) {
      result.push("font:"+i+":"+xlt.newfonts[i]);
      }

   for (i=1;i<xlt.newlayouts.length;i++) {
      result.push("layout:"+i+":"+xlt.newlayouts[i]);
      }

   for (i=1;i<xlt.newvalueformats.length;i++) {
      result.push("valueformat:"+i+":"+SocialCalc.encodeForSave(xlt.newvalueformats[i]));
      }

   for (i=0; i<xlt.namesorder.length; i++) {
      name = xlt.namesorder[i];
      result.push("name:"+SocialCalc.encodeForSave(name).toUpperCase()+":"+
                   SocialCalc.encodeForSave(sheetobj.names[name].desc)+":"+
                   SocialCalc.encodeForSave(sheetobj.names[name].definition));
      }

   if (range) {
      result.push("copiedfrom:"+SocialCalc.crToCoord(cr1.col, cr1.row)+":"+
                  SocialCalc.crToCoord(cr2.col, cr2.row));
      }

   result.push(""); // one extra to get extra \n

   delete sheetobj.xlt; // clean up

   return result.join("\n");
   }

//
// line = SocialCalc.CellToString(sheet, cell)
//

SocialCalc.CellToString = function(sheet, cell) {

   var cell, line, value, formula, t, r, b, l, xlt;

   line = "";

   if (!cell) return line;

   value = SocialCalc.encodeForSave(cell.datavalue);
   if (cell.datatype=="v") {
      if (cell.valuetype=="n") line += ":v:"+value;
      else line += ":vt:"+cell.valuetype+":"+value;
      }
   else if (cell.datatype=="t") {
      if (cell.valuetype==SocialCalc.Constants.textdatadefaulttype)
         line += ":t:"+value;
      else line += ":vt:"+cell.valuetype+":"+value;
      }
   else {
      formula = SocialCalc.encodeForSave(cell.formula);
      if (cell.datatype=="f") {
         line += ":vtf:"+cell.valuetype+":"+value+":"+formula;
         }
      else if (cell.datatype=="c") {
         line += ":vtc:"+cell.valuetype+":"+value+":"+formula;
         }
      }
   if (cell.readonly) {
      line += ":ro:yes";
      }
   if (cell.errors) {
      line += ":e:"+SocialCalc.encodeForSave(cell.errors);
      }
   t = cell.bt || "";
   r = cell.br || "";
   b = cell.bb || "";
   l = cell.bl || "";

   if (sheet.xlt) { // if have canonical save info
      xlt = sheet.xlt;
      if (t || r || b || l)
      line += ":b:"+xlt.borderstylesxlat[t||0]+":"+xlt.borderstylesxlat[r||0]+":"+xlt.borderstylesxlat[b||0]+":"+xlt.borderstylesxlat[l||0];
      if (cell.layout) line += ":l:"+xlt.layoutsxlat[cell.layout];
      if (cell.font) line += ":f:"+xlt.fontsxlat[cell.font];
      if (cell.color) line += ":c:"+xlt.colorsxlat[cell.color];
      if (cell.bgcolor) line += ":bg:"+xlt.colorsxlat[cell.bgcolor];
      if (cell.cellformat) line += ":cf:"+xlt.cellformatsxlat[cell.cellformat];
      if (cell.textvalueformat) line += ":tvf:"+xlt.valueformatsxlat[cell.textvalueformat];
      if (cell.nontextvalueformat) line += ":ntvf:"+xlt.valueformatsxlat[cell.nontextvalueformat];
      }
   else {
      if (t || r || b || l)
      line += ":b:"+t+":"+r+":"+b+":"+l;
      if (cell.layout) line += ":l:"+cell.layout;
      if (cell.font) line += ":f:"+cell.font;
      if (cell.color) line += ":c:"+cell.color;
      if (cell.bgcolor) line += ":bg:"+cell.bgcolor;
      if (cell.cellformat) line += ":cf:"+cell.cellformat;
      if (cell.textvalueformat) line += ":tvf:"+cell.textvalueformat;
      if (cell.nontextvalueformat) line += ":ntvf:"+cell.nontextvalueformat;
      }
   if (cell.colspan) line += ":colspan:"+cell.colspan;
   if (cell.rowspan) line += ":rowspan:"+cell.rowspan;
   if (cell.cssc) line += ":cssc:"+cell.cssc;
   if (cell.csss) line += ":csss:"+SocialCalc.encodeForSave(cell.csss);
   if (cell.mod) line += ":mod:"+cell.mod;
   if (cell.comment) line += ":comment:"+SocialCalc.encodeForSave(cell.comment);

   return line;

   }

//
// SocialCalc.CanonicalizeSheet(sheetobj, full)
//
// Goes through the sheet and fills in sheetobj.xlt with the following:
//
//   .maxrow, .maxcol - lastrow and lastcol are as small as possible
//   .newlayouts - new version of sheetobj.layouts without unused ones and all in ascending order
//   .layoutsxlat - maps old layouts index to new one
//   same ".new" and ".xlat" for fonts, colors, borderstyles, cell and value formats
//   .namesorder - array with names sorted
//
// If full or SocialCalc.Constants.doCanonicalizeSheet are not true, then the values will leave things unchanged (to save time, etc.)
//
// sheetobj.xlt should be deleted when you are finished using it
//

SocialCalc.CanonicalizeSheet = function(sheetobj, full) {

   var l, coord, cr, cell, filled, an, a, newa, newxlat, used, ahash, i, v;
   var maxrow = 0;
   var maxcol = 0;
   var alist = ["borderstyle", "cellformat", "color", "font", "layout", "valueformat"];

   var xlt = {};

   xlt.namesorder = []; // always return a sorted list
   for (a in sheetobj.names) {
      xlt.namesorder.push(a);
      }
   xlt.namesorder.sort();

   if (!SocialCalc.Constants.doCanonicalizeSheet || !full) { // return make-no-changes values if not wanted
      for (an=0; an<alist.length; an++) {
         a = alist[an];
         xlt["new"+a+"s"] = sheetobj[a+"s"];
         l = sheetobj[a+"s"].length;
         newxlat = new Array(l);
         newxlat[0] = "";
         for (i=1; i<l; i++) {
            newxlat[i] = i;
            }
         xlt[a+"sxlat"] = newxlat;
         }

      xlt.maxrow = sheetobj.attribs.lastrow;
      xlt.maxcol = sheetobj.attribs.lastcol;

      sheetobj.xlt = xlt;

      return;
      }

   for (an=0; an<alist.length; an++) {
      a = alist[an];
      xlt[a+"sUsed"] = {};
      }

   var colorsUsed = xlt.colorsUsed;
   var borderstylesUsed = xlt.borderstylesUsed;
   var fontsUsed = xlt.fontsUsed;
   var layoutsUsed = xlt.layoutsUsed;
   var cellformatsUsed = xlt.cellformatsUsed;
   var valueformatsUsed = xlt.valueformatsUsed;

   for (coord in sheetobj.cells) { // check all cells to see which values are used
      cr = SocialCalc.coordToCr(coord);
      cell = sheetobj.cells[coord];
      filled = false;

      if (cell.valuetype && cell.valuetype!="b") filled = true;

      if (cell.color) {
         colorsUsed[cell.color] = 1;
         filled = true;
         }

      if (cell.bgcolor) {
         colorsUsed[cell.bgcolor] = 1;
         filled = true;
         }

      if (cell.bt) {
         borderstylesUsed[cell.bt] = 1;
         filled = true;
         }
      if (cell.br) {
         borderstylesUsed[cell.br] = 1;
         filled = true;
         }
      if (cell.bb) {
         borderstylesUsed[cell.bb] = 1;
         filled = true;
         }
      if (cell.bl) {
         borderstylesUsed[cell.bl] = 1;
         filled = true;
         }

      if (cell.layout) {
         layoutsUsed[cell.layout] = 1;
         filled = true;
         }

      if (cell.font) {
         fontsUsed[cell.font] = 1;
         filled = true;
         }

      if (cell.cellformat) {
         cellformatsUsed[cell.cellformat] = 1;
         filled = true;
         }

      if (cell.textvalueformat) {
         valueformatsUsed[cell.textvalueformat] = 1;
         filled = true;
         }

      if (cell.nontextvalueformat) {
         valueformatsUsed[cell.nontextvalueformat] = 1;
         filled = true;
         }

      if (filled) {
         if (cr.row > maxrow) maxrow = cr.row;
         if (cr.col > maxcol) maxcol = cr.col;
         }
      }

   for (i=0; i<SocialCalc.sheetfieldsxlat.length; i++) { // do sheet values, too
      v = sheetobj.attribs[SocialCalc.sheetfieldsxlat[i]];
      if (v) {
         xlt[SocialCalc.sheetfieldsxlatxlt[i]+"sUsed"][v] = 1;
         }
      }

   a = {"height": 1, "hide": 1}; // look at explicit row settings
   for (v in a) {
      for (cr in sheetobj.rowattribs[v]) {
         if (cr > maxrow) maxrow = cr;
         }
      }
   a = {"hide": 1, "width": 1}; // look at explicit col settings
   for (v in a) {
      for (coord in sheetobj.colattribs[v]) {
         cr = SocialCalc.coordToCr(coord+"1");
         if (cr.col > maxcol) maxcol = cr.col;
         }
      }

   for (an=0; an<alist.length; an++) { // go through the attribs we want
      a = alist[an];

      newa = [];
      used = xlt[a+"sUsed"];
      for (v in used) {
         newa.push(sheetobj[a+"s"][v]);
         }
      newa.sort();
      newa.unshift("");

      newxlat = [""];
      ahash = sheetobj[a+"hash"];

      for (i=1; i<newa.length; i++) {
         newxlat[ahash[newa[i]]] = i;
         }

      xlt[a+"sxlat"] = newxlat;
      xlt["new"+a+"s"] = newa;

      }

   xlt.maxrow = maxrow || 1;
   xlt.maxcol = maxcol || 1;

   sheetobj.xlt = xlt; // leave for use by caller

   }

//
// result = SocialCalc.EncodeCellAttributes(sheet, coord)
//
// Returns the cell's attributes in an object, each in the following form:
//
//    attribname: {def: true/false, val: full-value}
//

SocialCalc.EncodeCellAttributes = function(sheet, coord) {

   var value, i, b, bb;
   var result = {};

   var InitAttrib = function(name) {
      result[name] = {def: true, val: ""};
      }

   var InitAttribs = function(namelist) {
      for (var i=0; i<namelist.length; i++) {
         InitAttrib(namelist[i]);
         }
      }

   var SetAttrib = function(name, v) {
      result[name].def = false;
      result[name].val = v || "";
      }

   var SetAttribStar = function(name, v) {
      if (v=="*") return;
      result[name].def = false;
      result[name].val = v;
      }

   var cell = sheet.GetAssuredCell(coord);

   // cellformat: alignhoriz

   InitAttrib("alignhoriz");
   if (cell.cellformat) {
      SetAttrib("alignhoriz", sheet.cellformats[cell.cellformat]);
      }

   // layout: alignvert, padtop, padright, padbottom, padleft

   InitAttribs(["alignvert", "padtop", "padright", "padbottom", "padleft"]);
   if (cell.layout) {
      parts = sheet.layouts[cell.layout].match(/^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);/);
      SetAttribStar("padtop", parts[1]);
      SetAttribStar("padright", parts[2]);
      SetAttribStar("padbottom", parts[3]);
      SetAttribStar("padleft", parts[4]);
      SetAttribStar("alignvert", parts[5]);
      }

   // font: fontfamily, fontlook, fontsize

   InitAttribs(["fontfamily", "fontlook", "fontsize"]);
   if (cell.font) {
      parts = sheet.fonts[cell.font].match(/^(\*|\S+? \S+?) (\S+?) (\S.*)$/);
      SetAttribStar("fontfamily", parts[3]);
      SetAttribStar("fontsize", parts[2]);
      SetAttribStar("fontlook", parts[1]);
      }

   // color: textcolor

   InitAttrib("textcolor");
   if (cell.color) {
      SetAttrib("textcolor", sheet.colors[cell.color]);
      }

   // bgcolor: bgcolor

   InitAttrib("bgcolor");
   if (cell.bgcolor) {
      SetAttrib("bgcolor", sheet.colors[cell.bgcolor]);
      }

   // formatting: numberformat, textformat

   InitAttribs(["numberformat", "textformat"]);
   if (cell.nontextvalueformat) {
      SetAttrib("numberformat", sheet.valueformats[cell.nontextvalueformat]);
      }
   if (cell.textvalueformat) {
      SetAttrib("textformat", sheet.valueformats[cell.textvalueformat]);
      }

   // merges: colspan, rowspan

   InitAttribs(["colspan", "rowspan"]);
   SetAttrib("colspan", cell.colspan || 1);
   SetAttrib("rowspan", cell.rowspan || 1);

   // borders: bXthickness, bXstyle, bXcolor for X = t, r, b, and l

   for (i=0; i<4; i++) {
      b = "trbl".charAt(i);
      bb = "b"+b;
      InitAttrib(bb);
      SetAttrib(bb, cell[bb] ? sheet.borderstyles[cell[bb]] : "");
      InitAttrib(bb+"thickness");
      InitAttrib(bb+"style");
      InitAttrib(bb+"color");
      if (cell[bb]) {
         parts = sheet.borderstyles[cell[bb]].match(/(\S+)\s+(\S+)\s+(\S.+)/);
         SetAttrib(bb+"thickness", parts[1]);
         SetAttrib(bb+"style", parts[2]);
         SetAttrib(bb+"color", parts[3]);
         }
      }
 
   // misc: cssc, csss, mod

   InitAttribs(["cssc", "csss", "mod"]);
   SetAttrib("cssc", cell.cssc || ""); 
   SetAttrib("csss", cell.csss || "");
   SetAttrib("mod", cell.mod || "n");

   return result;

   }

//
// result = SocialCalc.EncodeSheetAttributes(sheet)
//
// Returns the sheet's attributes in an object, each in the following form:
//
//    attribname: {def: true/false, val: full-value}
//

SocialCalc.EncodeSheetAttributes = function(sheet) {

   var value;
   var attribs = sheet.attribs;
   var result = {};

   var InitAttrib = function(name) {
      result[name] = {def: true, val: ""};
      }

   var InitAttribs = function(namelist) {
      for (var i=0; i<namelist.length; i++) {
         InitAttrib(namelist[i]);
         }
      }

   var SetAttrib = function(name, v) {
      result[name].def = false;
      result[name].val = v || value;
      }

   var SetAttribStar = function(name, v) {
      if (v=="*") return;
      result[name].def = false;
      result[name].val = v;
      }

   // sizes: colwidth, rowheight

   InitAttrib("colwidth");
   if (attribs.defaultcolwidth) {
      SetAttrib("colwidth", attribs.defaultcolwidth);
      }

   InitAttrib("rowheight");
   if (attribs.rowheight) {
      SetAttrib("rowheight", attribs.defaultrowheight);
      }

   // cellformat: textalignhoriz, numberalignhoriz

   InitAttrib("textalignhoriz");
   if (attribs.defaulttextformat) {
      SetAttrib("textalignhoriz", sheet.cellformats[attribs.defaulttextformat]);
      }

   InitAttrib("numberalignhoriz");
   if (attribs.defaultnontextformat) {
      SetAttrib("numberalignhoriz", sheet.cellformats[attribs.defaultnontextformat]);
      }

   // layout: alignvert, padtop, padright, padbottom, padleft

   InitAttribs(["alignvert", "padtop", "padright", "padbottom", "padleft"]);
   if (attribs.defaultlayout) {
      parts = sheet.layouts[attribs.defaultlayout].match(/^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);/);
      SetAttribStar("padtop", parts[1]);
      SetAttribStar("padright", parts[2]);
      SetAttribStar("padbottom", parts[3]);
      SetAttribStar("padleft", parts[4]);
      SetAttribStar("alignvert", parts[5]);
      }

   // font: fontfamily, fontlook, fontsize

   InitAttribs(["fontfamily", "fontlook", "fontsize"]);
   if (attribs.defaultfont) {
      parts = sheet.fonts[attribs.defaultfont].match(/^(\*|\S+? \S+?) (\S+?) (\S.*)$/);
      SetAttribStar("fontfamily", parts[3]);
      SetAttribStar("fontsize", parts[2]);
      SetAttribStar("fontlook", parts[1]);
      }

   // color: textcolor

   InitAttrib("textcolor");
   if (attribs.defaultcolor) {
      SetAttrib("textcolor", sheet.colors[attribs.defaultcolor]);
      }

   // bgcolor: bgcolor

   InitAttrib("bgcolor");
   if (attribs.defaultbgcolor) {
      SetAttrib("bgcolor", sheet.colors[attribs.defaultbgcolor]);
      }

   // formatting: numberformat, textformat

   InitAttribs(["numberformat", "textformat"]);
   if (attribs.defaultnontextvalueformat) {
      SetAttrib("numberformat", sheet.valueformats[attribs.defaultnontextvalueformat]);
      }
   if (attribs.defaulttextvalueformat) {
      SetAttrib("textformat", sheet.valueformats[attribs.defaulttextvalueformat]);
      }

   // recalc: recalc

   InitAttrib("recalc");
   if (attribs.recalc) {
      SetAttrib("recalc", attribs.recalc);
      }

   // usermaxcol, usermaxrow
   InitAttrib("usermaxcol");
   if (attribs.usermaxcol) {
      SetAttrib("usermaxcol", attribs.usermaxcol);
      }
   InitAttrib("usermaxrow");
   if (attribs.usermaxrow) {
      SetAttrib("usermaxrow", attribs.usermaxrow);
      }

   return result;

   }

//
// cmdstr = SocialCalc.DecodeCellAttributes(sheet, coord, attribs, range)
//
// Takes cell attributes in an object, each in the following form:
//
//    attribname: {def: true/false, val: full-value}
//
// and returns the sheet commands to make the actual attributes correspond.
// Returns a non-null string if any commands are to be executed, null otherwise.
//
// If range is provided, the commands are executed on the whole range.
//

SocialCalc.DecodeCellAttributes = function(sheet, coord, newattribs, range) {

   var value, b, bb;

   var cell = sheet.GetAssuredCell(coord);

   var changed = false;

   var CheckChanges = function(attribname, oldval, cmdname) {
      var val;
      if (newattribs[attribname]) {
         if (newattribs[attribname].def) {
            val = "";
            }
         else {
            val = newattribs[attribname].val;
            }
         if (val != (oldval || "")) {
            DoCmd(cmdname+" "+val);
            }
         }
      }

   var cmdstr = "";

   var DoCmd = function(str) {
      if (cmdstr) cmdstr += "\n";
      cmdstr += "set "+(range || coord)+" "+str;
      changed = true;
      }

   // cellformat: alignhoriz

   CheckChanges("alignhoriz", sheet.cellformats[cell.cellformat], "cellformat");

   // layout: alignvert, padtop, padright, padbottom, padleft

   if (!newattribs.alignvert.def || !newattribs.padtop.def || !newattribs.padright.def ||
       !newattribs.padbottom.def || !newattribs.padleft.def) {
      value = "padding:" +
         (newattribs.padtop.def ? "* " : newattribs.padtop.val + " ") +
         (newattribs.padright.def ? "* " : newattribs.padright.val + " ") +
         (newattribs.padbottom.def ? "* " : newattribs.padbottom.val + " ") +
         (newattribs.padleft.def ? "*" : newattribs.padleft.val) +
         ";vertical-align:" +
         (newattribs.alignvert.def ? "*;" : newattribs.alignvert.val+";");
      }
   else {
      value = "";
      }

   if (value != (sheet.layouts[cell.layout] || "")) {
      DoCmd("layout "+value);
      }

   // font: fontfamily, fontlook, fontsize

   if (!newattribs.fontlook.def || !newattribs.fontsize.def || !newattribs.fontfamily.def) {
      value =
         (newattribs.fontlook.def ? "* " : newattribs.fontlook.val + " ") +
         (newattribs.fontsize.def ? "* " : newattribs.fontsize.val + " ") +
         (newattribs.fontfamily.def ? "*" : newattribs.fontfamily.val);
      }
   else {
      value = "";
      }

   if (value != (sheet.fonts[cell.font] || "")) {
      DoCmd("font "+value);
      }

   // color: textcolor

   CheckChanges("textcolor", sheet.colors[cell.color], "color");

   // bgcolor: bgcolor

   CheckChanges("bgcolor", sheet.colors[cell.bgcolor], "bgcolor");

   // formatting: numberformat, textformat

   CheckChanges("numberformat", sheet.valueformats[cell.nontextvalueformat], "nontextvalueformat");

   CheckChanges("textformat", sheet.valueformats[cell.textvalueformat], "textvalueformat");

   // merges: colspan, rowspan - NOT HANDLED: IGNORED!

   // borders: bX for X = t, r, b, and l; bXthickness, bXstyle, bXcolor ignored

   for (i=0; i<4; i++) {
      b = "trbl".charAt(i);
      bb = "b"+b;
      CheckChanges(bb, sheet.borderstyles[cell[bb]], bb);
      }

   // misc: cssc, csss, mod

   CheckChanges("cssc", cell.cssc, "cssc");

   CheckChanges("csss", cell.csss, "csss");

   if (newattribs.mod) {
      if (newattribs.mod.def) {
         value = "n";
         }
      else {
         value = newattribs.mod.val;
         }
      if (value != (cell.mod || "n")) {
         if (value=="n") value = ""; // restrict to "y" and "" normally
         DoCmd("mod "+value);
         }
      }

   // if any changes return command(s)

   if (changed) {
       return cmdstr;
      }
   else {
      return null;
      }

   }


//
// changed = SocialCalc.DecodeSheetAttributes(sheet, newattribs)
//
// Takes sheet attributes in an object, each in the following form:
//
//    attribname: {def: true/false, val: full-value}
//
// and returns the sheet commands to make the actual attributes correspond.
// Returns a non-null string if any commands were executed, null otherwise.
//

SocialCalc.DecodeSheetAttributes = function(sheet, newattribs) {

   var value;
   var attribs = sheet.attribs;
   var changed = false;

   var CheckChanges = function(attribname, oldval, cmdname) {
      var val;
      if (newattribs[attribname]) {
         if (newattribs[attribname].def) {
            val = "";
            }
         else {
            val = newattribs[attribname].val;
            }
         if (val != (oldval || "")) {
            DoCmd(cmdname+" "+val);
            }
         }
      }

   var cmdstr = "";

   var DoCmd = function(str) {
      if (cmdstr) cmdstr += "\n";
      cmdstr += "set sheet "+str;
      changed = true;
      }

   // sizes: colwidth, rowheight

   CheckChanges("colwidth", attribs.defaultcolwidth, "defaultcolwidth");

   CheckChanges("rowheight", attribs.defaultrowheight, "defaultrowheight");

   // cellformat: textalignhoriz, numberalignhoriz

   CheckChanges("textalignhoriz", sheet.cellformats[attribs.defaulttextformat], "defaulttextformat");

   CheckChanges("numberalignhoriz", sheet.cellformats[attribs.defaultnontextformat], "defaultnontextformat");

   // layout: alignvert, padtop, padright, padbottom, padleft

   if (!newattribs.alignvert.def || !newattribs.padtop.def || !newattribs.padright.def ||
       !newattribs.padbottom.def || !newattribs.padleft.def) {
      value = "padding:" +
         (newattribs.padtop.def ? "* " : newattribs.padtop.val + " ") +
         (newattribs.padright.def ? "* " : newattribs.padright.val + " ") +
         (newattribs.padbottom.def ? "* " : newattribs.padbottom.val + " ") +
         (newattribs.padleft.def ? "*" : newattribs.padleft.val) +
         ";vertical-align:" +
         (newattribs.alignvert.def ? "*;" : newattribs.alignvert.val+";");
      }
   else {
      value = "";
      }

   if (value != (sheet.layouts[attribs.defaultlayout] || "")) {
      DoCmd("defaultlayout "+value);
      }

   // font: fontfamily, fontlook, fontsize

   if (!newattribs.fontlook.def || !newattribs.fontsize.def || !newattribs.fontfamily.def) {
      value =
         (newattribs.fontlook.def ? "* " : newattribs.fontlook.val + " ") +
         (newattribs.fontsize.def ? "* " : newattribs.fontsize.val + " ") +
         (newattribs.fontfamily.def ? "*" : newattribs.fontfamily.val);
      }
   else {
      value = "";
      }

   if (value != (sheet.fonts[attribs.defaultfont] || "")) {
      DoCmd("defaultfont "+value);
      }

   // color: textcolor

   CheckChanges("textcolor", sheet.colors[attribs.defaultcolor], "defaultcolor");

   // bgcolor: bgcolor

   CheckChanges("bgcolor", sheet.colors[attribs.defaultbgcolor], "defaultbgcolor");

   // formatting: numberformat, textformat

   CheckChanges("numberformat", sheet.valueformats[attribs.defaultnontextvalueformat], "defaultnontextvalueformat");

   CheckChanges("textformat", sheet.valueformats[attribs.defaulttextvalueformat], "defaulttextvalueformat");

   // recalc: recalc

   CheckChanges("recalc", sheet.attribs.recalc, "recalc");

   // usermaxcol, usermaxrow

   CheckChanges("usermaxcol", sheet.attribs.usermaxcol, "usermaxcol");
   CheckChanges("usermaxrow", sheet.attribs.usermaxrow, "usermaxrow");

   // if any changes return command(s)

   if (changed) {
       return cmdstr;
      }
   else {
      return null;
      }

   }

// *************************************
//
// Sheet command routines
//
// *************************************

//
// SocialCalc.SheetCommandInfo - object with information used during command execution
//

SocialCalc.SheetCommandInfo = function(sheetobj) {

   this.sheetobj = sheetobj; // sheet being operated on
   this.timerobj = null; // used for timeslicing
   this.firsttimerdelay = 50; // wait before starting cmds (for Chrome - to give time to update)
   this.timerdelay = 1; // wait between slices
   this.maxtimeslice = 100; // do another slice after this many milliseconds
   this.saveundo = false; // arg for ExecuteSheetCommand

   this.CmdExtensionCallbacks = {}; // for startcmdextension, in form: cmdname, {func:function(cmdname, data, sheet, SocialCalc.Parse object, saveundo), data:whatever}

//   statuscallback: null, // called during execution - obsolete: use sheet obj's
//   statuscallbackparams: null

   };

//
// SocialCalc.ScheduleSheetCommands
//
// statuscallback is called at the beginning (cmdstart) and end (cmdend).
//

SocialCalc.ScheduleSheetCommands = function(sheet, cmdstr, saveundo) {

   var sci = sheet.sci;

   var parseobj = new SocialCalc.Parse(cmdstr);

   if (sci.sheetobj.statuscallback) { // notify others if requested
      sheet.statuscallback(sci, "cmdstart", "", sci.sheetobj.statuscallbackparams);
      }

   if (saveundo) {
      sci.sheetobj.changes.PushChange(""); // add a step to undo stack
      }

   sci.timerobj = window.setTimeout(function() {
      SocialCalc.SheetCommandsTimerRoutine(sci, parseobj, saveundo);
   }, sci.firsttimerdelay);

   }

SocialCalc.SheetCommandsTimerRoutine = function(sci, parseobj, saveundo) {

   var errortext;
   var starttime = new Date();

   sci.timerobj = null;

   while (!parseobj.EOF()) { // go through all commands (separated by newlines)

      errortext = SocialCalc.ExecuteSheetCommand(sci.sheetobj, parseobj, saveundo);
      if (errortext) alert(errortext);

      parseobj.NextLine();

      if (((new Date()) - starttime) >= sci.maxtimeslice) { // if taking too long, give up CPU for a while
         sci.timerobj = window.setTimeout(function() {
            SocialCalc.SheetCommandsTimerRoutine(sci, parseobj, saveundo);
         }, sci.timerdelay);
         return;
         }
      }

   if (sci.sheetobj.statuscallback) { // notify others if requested
      sci.sheetobj.statuscallback(sci, "cmdend", "", sci.sheetobj.statuscallbackparams);
      }

   }

//
// errortext = SocialCalc.ExecuteSheetCommand(sheet, cmd, saveundo)
//
// cmd is a SocialCalc.Parse object.
//
// Executes commands that modify the sheet data.
// Sets sheet "needsrecalc" as needed.
// Sets sheet "changedrendervalues" as needed.
//
// The cmd string may be multiple commands, separated by newlines. In that case
// only one "step" is put on the undo stack representing all the commands.
// Note that because of this, in "set A1 text ..." and "set A1 comment ..." text is
// treated as encoded (newline => \n, \ => \b, : => \c).
//
// The commands are in the forms:
//
//    set sheet attributename value (plus lastcol and lastrow)
//    set 22 attributename value
//    set B attributename value
//    set A1 attributename value1 value2... (see each attribute in code for details)
//    set A1:B5 attributename value1 value2...
//    erase/copy/cut/paste/fillright/filldown A1:B5 all/formulas/format
//    loadclipboard save-encoded-clipboard-data
//    clearclipboard
//    merge C3:F3
//    unmerge C3
//    insertcol/insertrow C5
//    deletecol/deleterow C5:E7
//    movepaste/moveinsert A1:B5 A8 all/formulas/format (if insert, destination must be in same rows or columns or else paste done)
//    sort cr1:cr2 col1 up/down col2 up/down col3 up/down
//    name define NAME definition
//    name desc NAME description
//    name delete NAME
//    recalc
//    redisplay
//    changedrendervalues
//    startcmdextension extension rest-of-command
//
// If saveundo is true, then undo information is saved in sheet.changes.
//

SocialCalc.ExecuteSheetCommand = function(sheet, cmd, saveundo) {

   var cmdstr, cmd1, rest, what, attrib, num, pos, pos2, errortext, undostart, val;
   var cr1, cr2, col, row, cr, cell, newcell;
   var fillright, rowstart, colstart, crbase, rowoffset, coloffset, basecell;
   var clipsheet, cliprange, numcols, numrows, attribtable;
   var colend, rowend, newcolstart, newrowstart, newcolend, newrowend, rownext, colnext, colthis, cellnext;
   var lastrow, lastcol, rowbefore, colbefore, oldformula, oldcr;
   var cols, dirs, lastsortcol, i, sortlist, sortcells, sortvalues, sorttypes;
   var sortfunction, slen, valtype, originalrow, sortedcr;
   var name, v1, v2;
   var cmdextension;

   var attribs = sheet.attribs;
   var changes = sheet.changes;
   var cellProperties = SocialCalc.CellProperties;
   var scc = SocialCalc.Constants;

   var ParseRange =
      function() {
         var prange = SocialCalc.ParseRange(what);
         cr1 = prange.cr1;
         cr2 = prange.cr2;
         if (cr2.col > attribs.lastcol) attribs.lastcol = cr2.col;
         if (cr2.row > attribs.lastrow) attribs.lastrow = cr2.row;
         };

   errortext = "";

   cmdstr = cmd.RestOfStringNoMove();
   if (saveundo) {
      sheet.changes.AddDo(cmdstr);
      }

   cmd1 = cmd.NextToken();

   switch (cmd1) {

      case "set":
         what = cmd.NextToken();
         attrib = cmd.NextToken();
         rest = cmd.RestOfString();
         undostart = "set "+what+" "+attrib;

         if (what=="sheet") {
            sheet.renderneeded = true;
            switch (attrib) {
               case "defaultcolwidth":
                  if (saveundo) changes.AddUndo(undostart, attribs[attrib]);
                  attribs[attrib] = rest;
                  break;
               case "defaultcolor":
               case "defaultbgcolor":
                  if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("color", attribs[attrib]));
                  attribs[attrib] = sheet.GetStyleNum("color", rest);
                  break;
               case "defaultlayout":
                  if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("layout", attribs[attrib]));
                  attribs[attrib] = sheet.GetStyleNum("layout", rest);
                  break;
               case "defaultfont":
                  if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("font", attribs[attrib]));
                  if (rest=="* * *") rest = ""; // all default
                  attribs[attrib] = sheet.GetStyleNum("font", rest);
                  break;
               case "defaulttextformat":
               case "defaultnontextformat":
                  if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("cellformat", attribs[attrib]));
                  attribs[attrib] = sheet.GetStyleNum("cellformat", rest);
                  break;
               case "defaulttextvalueformat":
               case "defaultnontextvalueformat":
                  if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("valueformat", attribs[attrib]));
                  attribs[attrib] = sheet.GetStyleNum("valueformat", rest);
                  for (cr in sheet.cells) { // forget all cached display strings
                     delete sheet.cells[cr].displaystring;
                     }
                  break;
               case "lastcol":
               case "lastrow":
                  if (saveundo) changes.AddUndo(undostart, attribs[attrib]-0);
                  num = rest-0;
                  if (typeof num == "number") attribs[attrib] = num > 0 ? num : 1;
                  break;
               case "recalc":
                  if (saveundo) changes.AddUndo(undostart, attribs[attrib]);
                  if (rest == "off") {
                     attribs.recalc = rest; // manual recalc, not auto
                     }
                  else { // all values other than "off" mean "on"
                     delete attribs.recalc;
                     }
                  break;
               case "usermaxcol":
               case "usermaxrow":
                  if (saveundo) changes.AddUndo(undostart, attribs[attrib]-0);
                  num = rest-0;
                  if (typeof num == "number") attribs[attrib] = num > 0 ? num : 0;
                  break;
               default:
                  errortext = scc.s_escUnknownSheetCmd+cmdstr;
                  break;
               }
            }

         else if (/^[a-z]{1,2}(:[a-z]{1,2})?$/i.test(what)) { // col attributes
            sheet.renderneeded = true;
            what = what.toUpperCase();
            pos = what.indexOf(":");
            if (pos>=0) {
               cr1 = SocialCalc.coordToCr(what.substring(0,pos)+"1");
               cr2 = SocialCalc.coordToCr(what.substring(pos+1)+"1");
               }
            else {
               cr1 = SocialCalc.coordToCr(what+"1");
               cr2 = cr1;
               }
            for (col=cr1.col; col <= cr2.col; col++) {
               if (attrib=="width") {
                  cr = SocialCalc.rcColname(col);
                  if (saveundo) changes.AddUndo("set "+cr+" width", sheet.colattribs.width[cr]);
                  if (rest.length > 0 ) {
                     sheet.colattribs.width[cr] = rest;
                     }
                  else {
                     delete sheet.colattribs.width[cr];
                     }
                  }
               else if (attrib=="hide") {
                  sheet.hiddencolrow = "col";
                  cr = SocialCalc.rcColname(col);
                  if (saveundo) changes.AddUndo("set "+cr+" hide", sheet.colattribs.hide[cr]);
                  if (rest.length > 0) {
                     sheet.colattribs.hide[cr] = rest; 
                     }
                  else {
                     delete sheet.colattribs.hide[cr];
                     }
                  }
               }
            }

         else if (/^\d+(:\d+)?$/i.test(what)) { // row attributes
            sheet.renderneeded = true;
            what = what.toUpperCase();
            pos = what.indexOf(":");
            if (pos>=0) {
               cr1 = SocialCalc.coordToCr("A"+what.substring(0,pos));
               cr2 = SocialCalc.coordToCr("A"+what.substring(pos+1));
               }
            else {
               cr1 = SocialCalc.coordToCr("A"+what);
               cr2 = cr1;
               }
            for (row=cr1.row; row <= cr2.row; row++) {
               if (attrib=="height") {
                  if (saveundo) changes.AddUndo("set "+row+" height", sheet.rowattribs.height[row]);
                  if (rest.length > 0 ) {
                     sheet.rowattribs.height[row] = rest;
                     }
                  else {
                     delete sheet.rowattribs.height[row];
                     }
                  }
               else if (attrib=="hide") {
                  sheet.hiddencolrow = "row";
                  if (saveundo) changes.AddUndo("set "+row+" hide", sheet.rowattribs.hide[row]);
                  if (rest.length > 0) {
                     sheet.rowattribs.hide[row] = rest; 
                     }
                  else {
                     delete sheet.rowattribs.hide[row];
                     }
                  }
               }
            }

         else if (/^[a-z]{1,2}\d+(:[a-z]{1,2}\d+)?$/i.test(what)) { // cell attributes
            ParseRange();
            if (cr1.row!=cr2.row || cr1.col!=cr2.col || sheet.celldisplayneeded || sheet.renderneeded) { // not one cell
               sheet.renderneeded = true;
               sheet.celldisplayneeded = "";
               }
            else {
               sheet.celldisplayneeded = SocialCalc.crToCoord(cr1.col, cr1.row);
               }
            for (row=cr1.row; row <= cr2.row; row++) {
               for (col=cr1.col; col <= cr2.col; col++) {
                  cr = SocialCalc.crToCoord(col, row);
                  cell=sheet.GetAssuredCell(cr);
                  if (cell.readonly && attrib!="readonly") continue;
                  if (saveundo) changes.AddUndo("set "+cr+" all", sheet.CellToString(cell));
                  if (attrib=="value") { // set coord value type numeric-value
                     pos = rest.indexOf(" ");
                     cell.datavalue = rest.substring(pos+1)-0;
                     delete cell.errors;
                     cell.datatype = "v";
                     cell.valuetype = rest.substring(0,pos);
                     delete cell.displaystring;
                     delete cell.parseinfo;
                     attribs.needsrecalc = "yes";
                     }
                  else if (attrib=="text") { // set coord text type text-value
                     pos = rest.indexOf(" ");
                     cell.datavalue = SocialCalc.decodeFromSave(rest.substring(pos+1));
                     delete cell.errors;
                     cell.datatype = "t";
                     cell.valuetype = rest.substring(0,pos);
                     delete cell.displaystring;
                     delete cell.parseinfo;
                     attribs.needsrecalc = "yes";
                     }
                  else if (attrib=="formula") { // set coord formula formula-body-less-initial-=
                     cell.datavalue = 0; // until recalc
                     delete cell.errors;
                     cell.datatype = "f";
                     cell.valuetype = "e#N/A"; // until recalc
                     cell.formula = rest;
                     delete cell.displaystring;
                     delete cell.parseinfo;
                     attribs.needsrecalc = "yes";
                     }
                  else if (attrib=="constant") { // set coord constant type numeric-value source-text
                     pos = rest.indexOf(" ");
                     pos2 = rest.substring(pos+1).indexOf(" ");
                     cell.datavalue = rest.substring(pos+1,pos+1+pos2)-0;
                     cell.valuetype = rest.substring(0,pos);
                     if (cell.valuetype.charAt(0)=="e") { // error
                        cell.errors = cell.valuetype.substring(1);
                        }
                     else {
                        delete cell.errors;
                        }
                     cell.datatype = "c";
                     cell.formula = rest.substring(pos+pos2+2);
                     delete cell.displaystring;
                     delete cell.parseinfo;
                     attribs.needsrecalc = "yes";
                     }
                  else if (attrib=="empty") { // erase value
                     cell.datavalue = "";
                     delete cell.errors;
                     cell.datatype = null;
                     cell.formula = "";
                     cell.valuetype = "b";
                     delete cell.displaystring;
                     delete cell.parseinfo;
                     attribs.needsrecalc = "yes";
                     }
                  else if (attrib=="all") { // set coord all :this:val1:that:val2...
                     if (rest.length>0) {
                        cell = new SocialCalc.Cell(cr);
                        sheet.CellFromStringParts(cell, rest.split(":"), 1);
                        sheet.cells[cr] = cell;
                        }
                     else {
                        delete sheet.cells[cr];
                        }
                     attribs.needsrecalc = "yes";
                     }
                  else if (/^b[trbl]$/.test(attrib)) { // set coord bt 1px solid black
                     cell[attrib] = sheet.GetStyleNum("borderstyle", rest);
                     sheet.renderneeded = true; // affects more than just one cell
                     }
                  else if (attrib=="color" || attrib=="bgcolor") {
                     cell[attrib] = sheet.GetStyleNum("color", rest);
                     }
                  else if (attrib=="layout" || attrib=="cellformat") {
                     cell[attrib] = sheet.GetStyleNum(attrib, rest);
                     }
                  else if (attrib=="font") { // set coord font style weight size family
                     if (rest=="* * *") rest = "";
                     cell[attrib] = sheet.GetStyleNum("font", rest);
                     }
                  else if (attrib=="textvalueformat" || attrib=="nontextvalueformat") {
                     cell[attrib] = sheet.GetStyleNum("valueformat", rest);
                     delete cell.displaystring;
                     }
                  else if (attrib=="cssc") {
                     rest = rest.replace(/[^a-zA-Z0-9\-]/g, "");
                     cell.cssc = rest;
                     }
                  else if (attrib=="csss") {
                     rest = rest.replace(/\n/g, "");
                     cell.csss = rest;
                     }
                  else if (attrib=="mod") {
                     rest = rest.replace(/[^yY]/g, "").toLowerCase();
                     cell.mod = rest;
                     }
                  else if (attrib=="comment") {
                     cell.comment = SocialCalc.decodeFromSave(rest);
                     }
                  else if (attrib=="readonly") {
                     cell.readonly = rest.toLowerCase()=="yes";
                     }
                  else {
                     errortext = scc.s_escUnknownSetCoordCmd+cmdstr;
                     }
                  }
               }

            }
         break;

      case "merge":
         sheet.renderneeded = true;
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         ParseRange();
         cell=sheet.GetAssuredCell(cr1.coord);
         if (cell.readonly) break;
         if (saveundo) changes.AddUndo("unmerge "+cr1.coord);

         if (cr2.col > cr1.col) cell.colspan = cr2.col - cr1.col + 1;
         else delete cell.colspan;
         if (cr2.row > cr1.row) cell.rowspan = cr2.row - cr1.row + 1;
         else delete cell.rowspan;

         sheet.changedrendervalues = true;

         break;

      case "unmerge":
         sheet.renderneeded = true;
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         ParseRange();
         cell=sheet.GetAssuredCell(cr1.coord);
         if (cell.readonly) break;
         if (saveundo) changes.AddUndo("merge "+cr1.coord+":"+SocialCalc.crToCoord(cr1.col+(cell.colspan||1)-1, cr1.row+(cell.rowspan||1)-1));

         delete cell.colspan;
         delete cell.rowspan;

         sheet.changedrendervalues = true;

         break;

      case "erase":
      case "cut":
         sheet.renderneeded = true;
         sheet.changedrendervalues = true;
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         ParseRange();

         if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
         if (cmd1=="cut") { // save copy of whole thing before erasing
            if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
            SocialCalc.Clipboard.clipboard = SocialCalc.CreateSheetSave(sheet, what);
            }

         for (row = cr1.row; row <= cr2.row; row++) {
            for (col = cr1.col; col <= cr2.col; col++) {
               cr = SocialCalc.crToCoord(col, row);
               cell=sheet.GetAssuredCell(cr);
               if (cell.readonly) continue;
               if (saveundo) changes.AddUndo("set "+cr+" all", sheet.CellToString(cell));
               if (rest=="all") {
                  delete sheet.cells[cr];
                  }
               else if (rest == "formulas") {
                  cell.datavalue = "";
                  cell.datatype = null;
                  cell.formula = "";
                  cell.valuetype = "b";
                  delete cell.errors;
                  delete cell.displaystring;
                  delete cell.parseinfo;
                  if (cell.comment) { // comments are considered content for erasing
                     delete cell.comment;
                     }
                  }
               else if (rest == "formats") {
                  newcell = new SocialCalc.Cell(cr); // create a new cell without attributes
                  newcell.datavalue = cell.datavalue; // copy existing values
                  newcell.datatype = cell.datatype;
                  newcell.formula = cell.formula;
                  newcell.valuetype = cell.valuetype;
                  if (cell.comment) {
                     newcell.comment = cell.comment;
                     }
                  sheet.cells[cr] = newcell; // replace
                  }
               }
            }
         attribs.needsrecalc = "yes";
         break;

      case "fillright":
      case "filldown":
         sheet.renderneeded = true;
         sheet.changedrendervalues = true;
         if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         ParseRange();
         if (cmd1 == "fillright") {
            fillright = true;
            rowstart = cr1.row;
            colstart = cr1.col + 1;
            }
         else {
            fillright = false;
            rowstart = cr1.row + 1;
            colstart = cr1.col;
            }
         for (row = rowstart; row <= cr2.row; row++) {
            for (col = colstart; col <= cr2.col; col++) {
               cr = SocialCalc.crToCoord(col, row);
               cell=sheet.GetAssuredCell(cr);
               if (cell.readonly) continue;
               if (saveundo) changes.AddUndo("set "+cr+" all", sheet.CellToString(cell));
               if (fillright) {
                  crbase = SocialCalc.crToCoord(cr1.col, row);
                  coloffset = col - colstart + 1;
                  rowoffset = 0;
                  }
               else {
                  crbase = SocialCalc.crToCoord(col, cr1.row);
                  coloffset = 0;
                  rowoffset = row - rowstart + 1;
                  }
               basecell = sheet.GetAssuredCell(crbase);
               if (rest == "all" || rest == "formats") {
                  for (attrib in cellProperties) {
                     if (cellProperties[attrib] == 1) continue; // copy only format attributes
                     if (typeof basecell[attrib] === undefined || cellProperties[attrib] == 3) {
                        delete cell[attrib];
                        }
                     else {
                        cell[attrib] = basecell[attrib];
                        }
                     }
                  }
               if (rest == "all" || rest == "formulas") {
                  cell.datavalue = basecell.datavalue;
                  cell.datatype = basecell.datatype;            
                  cell.valuetype = basecell.valuetype;
                  if (cell.datatype == "f") { // offset relative coords, even in sheet references
                     cell.formula = SocialCalc.OffsetFormulaCoords(basecell.formula, coloffset, rowoffset);
                     }
                  else {
                     cell.formula = basecell.formula;
                     }
                  delete cell.parseinfo;
                  cell.errors = basecell.errors;
                  }
               delete cell.displaystring;
               }
            }

         attribs.needsrecalc = "yes";
         break;

      case "copy":
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
         SocialCalc.Clipboard.clipboard = SocialCalc.CreateSheetSave(sheet, what);
         break;

      case "loadclipboard":
         rest = cmd.RestOfString();
         if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
         SocialCalc.Clipboard.clipboard = SocialCalc.decodeFromSave(rest);
         break;

      case "clearclipboard":
         if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
         SocialCalc.Clipboard.clipboard = "";
         break;

      case "paste":
         sheet.renderneeded = true;
         sheet.changedrendervalues = true;
         if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         ParseRange();
         if (!SocialCalc.Clipboard.clipboard) {
            break;
            }
         clipsheet = new SocialCalc.Sheet(); // load clipboard contents as another sheet
         clipsheet.ParseSheetSave(SocialCalc.Clipboard.clipboard);
         cliprange = SocialCalc.ParseRange(clipsheet.copiedfrom);
         coloffset = cr1.col - cliprange.cr1.col; // get sizes, etc.
         rowoffset = cr1.row - cliprange.cr1.row;
         numcols = Math.max(cr2.col - cr1.col + 1, cliprange.cr2.col - cliprange.cr1.col + 1);
         numrows = Math.max(cr2.row - cr1.row + 1, cliprange.cr2.row - cliprange.cr1.row + 1);
         if (cr1.col+numcols-1 > attribs.lastcol) attribs.lastcol = cr1.col+numcols-1;
         if (cr1.row+numrows-1 > attribs.lastrow) attribs.lastrow = cr1.row+numrows-1;

         for (row = cr1.row; row < cr1.row+numrows; row++) {
            for (col = cr1.col; col < cr1.col+numcols; col++) {
               cr = SocialCalc.crToCoord(col, row);
               cell=sheet.GetAssuredCell(cr);
               if (cell.readonly) continue;
               if (saveundo) changes.AddUndo("set "+cr+" all", sheet.CellToString(cell));
               crbase = SocialCalc.crToCoord(
                  cliprange.cr1.col + ((col-cr1.col) % (cliprange.cr2.col - cliprange.cr1.col + 1)), 
                  cliprange.cr1.row + ((row-cr1.row) % (cliprange.cr2.row - cliprange.cr1.row + 1)));
               basecell = clipsheet.GetAssuredCell(crbase);
               if (rest == "all" || rest == "formats") {
                  for (attrib in cellProperties) {
                     if (cellProperties[attrib] == 1) continue; // copy only format attributes
                     if (typeof basecell[attrib] === undefined || cellProperties[attrib] == 3) {
                        delete cell[attrib];
                        }
                     else {
                        attribtable = SocialCalc.CellPropertiesTable[attrib];
                        if (attribtable && basecell[attrib]) { // table indexes to expand to strings since other sheet may have diff indexes
                           cell[attrib] = sheet.GetStyleNum(attribtable, clipsheet.GetStyleString(attribtable, basecell[attrib]));
                           }
                        else { // these are not table indexes
                           cell[attrib] = basecell[attrib];
                           }
                        }
                     }
                  }
               if (rest == "all" || rest == "formulas") {
                  cell.datavalue = basecell.datavalue;
                  cell.datatype = basecell.datatype;            
                  cell.valuetype = basecell.valuetype;
                  if (cell.datatype == "f") { // offset relative coords, even in sheet references
                     cell.formula = SocialCalc.OffsetFormulaCoords(basecell.formula, coloffset, rowoffset);
                     }
                  else {
                     cell.formula = basecell.formula;
                     }
                  delete cell.parseinfo;
                  cell.errors = basecell.errors;
                  if (basecell.comment) { // comments are pasted as part of content, though not filled, etc.
                     cell.comment = basecell.comment;
                     }
                  else if (cell.comment) {
                     delete cell.comment;
                     }
                  }
               delete cell.displaystring;
               }
            }

         attribs.needsrecalc = "yes";
         break;

      case "sort": // sort cr1:cr2 col1 up/down col2 up/down col3 up/down
         sheet.renderneeded = true;
         sheet.changedrendervalues = true;
         if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
         what = cmd.NextToken();
         ParseRange();
         cols = []; // get columns and sort directions (or "")
         dirs = [];
         lastsortcol = 0;
         for (i=0; i<=3; i++) {
            cols[i] = cmd.NextToken();
            dirs[i] = cmd.NextToken();
            if (cols[i]) lastsortcol = i;
            }

         sortcells = {}; // a copy of the data which will replace the original, but in the new order
         sortlist = []; // an array of 0, 1, ..., nrows-1 needed for sorting
         sortvalues = []; // values to be sorted corresponding to sortlist
         sorttypes = []; // basic types of the values

         for (row = cr1.row; row <= cr2.row; row++) { // fill in the sort info
            for (col = cr1.col; col <= cr2.col; col++) {
               cr = SocialCalc.crToCoord(col, row);
               cell=sheet.cells[cr];
               if (cell) { // only copy non-empty cells
                  sortcells[cr] = sheet.CellToString(cell);
                  if (saveundo) changes.AddUndo("set "+cr+" all", sortcells[cr]);
                  }
               else {
                  if (saveundo) changes.AddUndo("set "+cr+" all");
                  }
               }
            sortlist.push(sortlist.length);
            sortvalues.push([]);
            sorttypes.push([]);
            slast = sorttypes.length-1;
            for (i = 0; i <= lastsortcol; i++) {
               cr = cols[i] + row; // get cr on this row in sort col
               cell = sheet.GetAssuredCell(cr);
               val = cell.datavalue;
               valtype = cell.valuetype.charAt(0) || "b";
               if (valtype == "t") val = val.toLowerCase();
               sortvalues[slast].push(val);
               sorttypes[slast].push(valtype);
               }
            }

         sortfunction = function(a, b) { // a comparison function that can handle all the type variations
            var i, a1, b1, ta, cresult;
            for (i=0; i<=lastsortcol; i++) {
               if (dirs[i] == "up") { // handle sort direction
                  a1 = a; b1 = b;
                  }
               else {
                  a1 = b; b1 = a;
                  }
               ta = sorttypes[a1][i];
               tb = sorttypes[b1][i];
               if (ta == "t") { // numbers < text < errors, blank always last no matter what dir
                  if (tb == "t") {
                     a1 = sortvalues[a1][i];
                     b1 = sortvalues[b1][i];
                     cresult = a1 > b1 ? 1 : (a1 < b1 ? -1 : 0);
                     }
                  else if (tb == "n") {
                     cresult = 1;
                     }
                  else if (tb == "b") {
                     cresult = dirs[i] == "up" ? -1 : 1;
                     }
                  else if (tb == "e") {
                     cresult = -1;
                     }
                  }
               else if (ta == "n") {
                  if (tb == "t") {
                     cresult = -1;
                     }
                  else if (tb == "n") {
                     a1 = sortvalues[a1][i]-0; // force to numeric, just in case
                     b1 = sortvalues[b1][i]-0;
                     cresult = a1 > b1 ? 1 : (a1 < b1 ? -1 : 0);
                     }
                  else if (tb == "b") {
                     cresult = dirs[i] == "up" ? -1 : 1;
                     }
                  else if (tb == "e") {
                     cresult = -1;
                     }
                  }
               else if (ta == "e") {
                  if (tb == "e") {
                     a1 = sortvalues[a1][i];
                     b1 = sortvalues[b1][i];
                     cresult = a1 > b1 ? 1 : (a1 < b1 ? -1 : 0);
                     }
                  else if (tb == "b") {
                     cresult = dirs[i] == "up" ? -1 : 1;
                     }
                  else {
                     cresult = 1;
                     }
                  }
               else if (ta == "b") {
                  if (tb == "b") {
                     cresult = 0;
                     }
                  else {
                     cresult = dirs[i] == "up" ? 1 : -1;
                     }
                  }
               if (cresult) { // return if tested not equal, otherwise do next column
                  return cresult;
                  }
               }
            cresult = a > b ? 1 : (a < b ? -1 : 0); // equal - return position in original to maintain it
            return cresult;
            }

         sortlist.sort(sortfunction);

         for (row = cr1.row; row <= cr2.row; row++) { // copy original rows into sorted positions
            originalrow = sortlist[row-cr1.row]; // relative position where it was in original
            for (col = cr1.col; col <= cr2.col; col++) {
               cr = SocialCalc.crToCoord(col, row);
               sortedcr = SocialCalc.crToCoord(col, originalrow+cr1.row); // original cell to be put in new place
               if (sortcells[sortedcr]) {
                  cell = new SocialCalc.Cell(cr);
                  sheet.CellFromStringParts(cell, sortcells[sortedcr].split(":"), 1);
                  if (cell.datatype == "f") { // offset coord refs, even to ***relative*** coords in other sheets
                     cell.formula = SocialCalc.OffsetFormulaCoords(cell.formula, 0, (row-cr1.row)-originalrow);
                     }
                  sheet.cells[cr] = cell;
                  }
               else {
                  delete sheet.cells[cr];
                  }
               }
            }

         attribs.needsrecalc = "yes";
         break;

      case "insertcol":
      case "insertrow":
         sheet.renderneeded = true;
         sheet.changedrendervalues = true;
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         ParseRange();

         if (cmd1 == "insertcol") {
            coloffset = 1;
            colend = cr1.col;
            rowoffset = 0;
            rowend = 1;
            newcolstart = cr1.col;
            newcolend = cr1.col;
            newrowstart = 1;
            newrowend = attribs.lastrow;
            if (saveundo) changes.AddUndo("deletecol "+cr1.coord);
            }
         else {
            coloffset = 0;
            colend = 1;
            rowoffset = 1;
            rowend = cr1.row;
            newcolstart = 1;
            newcolend = attribs.lastcol;
            newrowstart = cr1.row;
            newrowend = cr1.row;
            if (saveundo) changes.AddUndo("deleterow "+cr1.coord);
            }

         for (row=attribs.lastrow; row >= rowend; row--) { // copy the cells forward
            for (col=attribs.lastcol; col >= colend; col--) {
               crbase = SocialCalc.crToCoord(col, row);
               cr = SocialCalc.crToCoord(col+coloffset, row+rowoffset);
               if (!sheet.cells[crbase]) { // copying empty cell
                  delete sheet.cells[cr]; // delete anything that may have been there
                  }
               else { // overwrite existing cell with moved contents
                  sheet.cells[cr] = sheet.cells[crbase];
                  }
               }
            }

         for (row=newrowstart; row <= newrowend; row++) { // fill the "new" empty cells
            for (col=newcolstart; col <= newcolend; col++) {
               cr = SocialCalc.crToCoord(col, row);
               cell = new SocialCalc.Cell(cr);
               sheet.cells[cr] = cell;
               crbase = SocialCalc.crToCoord(col-coloffset, row-rowoffset); // copy attribs of the one before (0 gives you A or 1)
               basecell = sheet.GetAssuredCell(crbase);
               for (attrib in cellProperties) {
                  if (cellProperties[attrib] == 2) { // copy only format attributes
                     cell[attrib] = basecell[attrib];
                     }
                  }
               }
            }

         for (cr in sheet.cells) { // update cell references to moved cells in calculated formulas
             cell = sheet.cells[cr];
             if (cell && cell.datatype == "f") {
                cell.formula = SocialCalc.AdjustFormulaCoords(cell.formula, cr1.col, coloffset, cr1.row, rowoffset);
                }
             if (cell) {
                delete cell.parseinfo;
                }
             }

         for (name in sheet.names) { // update cell references to moved cells in names
            if (sheet.names[name]) { // works with "A1", "A1:A20", and "=formula" forms
               v1 = sheet.names[name].definition;
               v2 = "";
               if (v1.charAt(0) == "=") {
                  v2 = "=";
                  v1 = v1.substring(1);
                  }
               sheet.names[name].definition = v2 +
                  SocialCalc.AdjustFormulaCoords(v1, cr1.col, coloffset, cr1.row, rowoffset);
               }
            }

         for (row = attribs.lastrow; row >= rowend && cmd1 == "insertrow"; row--) { // copy the row attributes forward
            rownext = row + rowoffset;
            for (attrib in sheet.rowattribs) {
               val = sheet.rowattribs[attrib][row];
               if (sheet.rowattribs[attrib][rownext] != val) { // make assignment only if different
                  if (val) {
                     sheet.rowattribs[attrib][rownext] = val;
                     }
                  else {
                     delete sheet.rowattribs[attrib][rownext];
                     }
                  }
               }
            }

         for (col = attribs.lastcol; col >= colend && cmd1 == "insertcol"; col--) { // copy the column attributes forward
            colthis = SocialCalc.rcColname(col);
            colnext = SocialCalc.rcColname(col + coloffset);
            for (attrib in sheet.colattribs) {
               val = sheet.colattribs[attrib][colthis];
               if (sheet.colattribs[attrib][colnext] != val) { // make assignment only if different
                  if (val) {
                     sheet.colattribs[attrib][colnext] = val;
                     }
                  else {
                     delete sheet.colattribs[attrib][colnext];
                     }
                  }
               }
            }

         attribs.lastcol += coloffset;
         attribs.lastrow += rowoffset;
         attribs.needsrecalc = "yes";
         break;

      case "deletecol":
      case "deleterow":
         sheet.renderneeded = true;
         sheet.changedrendervalues = true;
         what = cmd.NextToken();
         rest = cmd.RestOfString();
         lastcol = attribs.lastcol; // save old values since ParseRange sets...
         lastrow = attribs.lastrow;
         ParseRange();

         if (cmd1 == "deletecol") {
            coloffset = cr1.col - cr2.col - 1;
            rowoffset = 0;
            colstart = cr2.col + 1;
            rowstart = 1;
            }
         else {
            coloffset = 0;
            rowoffset = cr1.row - cr2.row - 1;
            colstart = 1;
            rowstart = cr2.row + 1;
            }

         for (row=rowstart; row <= lastrow - rowoffset; row++) { // check for readonly cells
            for (col=colstart; col <= lastcol - coloffset; col++) {
               cr = SocialCalc.crToCoord(col+coloffset, row+rowoffset);
               cell = sheet.cells[cr];
               if (cell && cell.readonly) return errortext; 
               }
            }

         for (row=rowstart; row <= lastrow - rowoffset; row++) { // copy the cells backwards - extra so no dup of last set
            for (col=colstart; col <= lastcol - coloffset; col++) {
               cr = SocialCalc.crToCoord(col+coloffset, row+rowoffset);
               if (saveundo && (row<rowstart-rowoffset || col<colstart	-coloffset)) { // save cells that are overwritten as undo info
                  cell = sheet.cells[cr];
                  if (!cell) { // empty cell
                     changes.AddUndo("erase "+cr+" all");
                     }
                  else {
                     changes.AddUndo("set "+cr+" all", sheet.CellToString(cell));
                     }
                  }
               crbase = SocialCalc.crToCoord(col, row);
               cell = sheet.cells[crbase];
               if (!cell) { // copying empty cell
                  delete sheet.cells[cr]; // delete anything that may have been there
                  }
               else { // overwrite existing cell with moved contents
                  sheet.cells[cr] = cell;
                  }
               }
            }

//!!! multiple deletes isn't setting #REF!; need to fix up #REF!'s on undo but only those!

         for (cr in sheet.cells) { // update cell references to moved cells in calculated formulas
             cell = sheet.cells[cr];
             if (cell) {
                if (cell.datatype == "f") {
                   oldformula = cell.formula;
                   cell.formula = SocialCalc.AdjustFormulaCoords(oldformula, cr1.col, coloffset, cr1.row, rowoffset);
                   if (cell.formula != oldformula) {
                      delete cell.parseinfo;
                      if (saveundo && cell.formula.indexOf("#REF!")!=-1) { // save old version only if removed coord
                         oldcr = SocialCalc.coordToCr(cr);
                         changes.AddUndo("set "+SocialCalc.rcColname(oldcr.col-coloffset)+(oldcr.row-rowoffset)+
                                         " formula "+oldformula);
                         }
                      }
                   }
                else {
                   delete cell.parseinfo;
                   }
                }
             }

         for (name in sheet.names) { // update cell references to moved cells in names
            if (sheet.names[name]) { // works with "A1", "A1:A20", and "=formula" forms
               v1 = sheet.names[name].definition;
               v2 = "";
               if (v1.charAt(0) == "=") {
                  v2 = "=";
                  v1 = v1.substring(1);
                  }
               sheet.names[name].definition = v2 +
                  SocialCalc.AdjustFormulaCoords(v1, cr1.col, coloffset, cr1.row, rowoffset);
               }
            }

         for (row = rowstart; row <= lastrow - rowoffset && cmd1 == "deleterow"; row++) { // copy the row attributes backwards
            rowbefore = row + rowoffset;
            for (attrib in sheet.rowattribs) {
               val = sheet.rowattribs[attrib][row];
               if (sheet.rowattribs[attrib][rowbefore] != val) { // make assignment only if different
                  if (saveundo) changes.AddUndo("set "+rowbefore+" "+attrib, sheet.rowattribs[attrib][rowbefore]);
                  if (val) {
                     sheet.rowattribs[attrib][rowbefore] = val;
                     }
                  else {
                     delete sheet.rowattribs[attrib][rowbefore];
                     }
                  }
               }
            }

         for (col = colstart; col <= lastcol - coloffset && cmd1 == "deletecol"; col++) { // copy the column attributes backwards
            colthis = SocialCalc.rcColname(col);
            colbefore = SocialCalc.rcColname(col + coloffset);
            for (attrib in sheet.colattribs) {
               val = sheet.colattribs[attrib][colthis];
               if (sheet.colattribs[attrib][colbefore] != val) { // make assignment only if different
                  if (saveundo) changes.AddUndo("set "+colbefore+" "+attrib, sheet.colattribs[attrib][colbefore]);
                  if (val) {
                     sheet.colattribs[attrib][colbefore] = val;
                     }
                  else {
                     delete sheet.colattribs[attrib][colbefore];
                     }
                  }
               }
            }

         if (saveundo) {
            if (cmd1 == "deletecol") {
               for (col=cr1.col; col<=cr2.col; col++) {
                  changes.AddUndo("insertcol "+SocialCalc.rcColname(col));
                  }
               }
            else {
               for (row=cr1.row; row<=cr2.row; row++) {
                  changes.AddUndo("insertrow "+row);
                  }
               }
            }

         if (cmd1 == "deletecol") {
            if (cr1.col <= lastcol) { // shrink sheet unless deleted phantom cols off the end
               if (cr2.col <= lastcol) {
                  attribs.lastcol += coloffset;
                  }
               else {
                  attribs.lastcol = cr1.col - 1;
                  }
               }
            }
         else {
            if (cr1.row <= lastrow) { // shrink sheet unless deleted phantom rows off the end
               if (cr2.row <= lastrow) {
                  attribs.lastrow += rowoffset;
                  }
               else {
                  attribs.lastrow = cr1.row - 1;
                  }
               }
            }
         attribs.needsrecalc = "yes";
         break;


      case "movepaste":
      case "moveinsert":

         var movingcells, dest, destcr, inserthoriz, insertvert, pushamount, movedto;

         sheet.renderneeded = true;
         sheet.changedrendervalues = true;
         if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
         what = cmd.NextToken();
         dest = cmd.NextToken();
         rest = cmd.RestOfString(); // rest is all/formulas/formats
         if (rest=="") rest = "all";

         ParseRange();

         destcr = SocialCalc.coordToCr(dest);

         coloffset = destcr.col - cr1.col;
         rowoffset = destcr.row - cr1.row;
         numcols = cr2.col - cr1.col + 1;
         numrows = cr2.row - cr1.row + 1;

         // get a copy of moving cells and erase from where they were

         movingcells = {};

         for (row = cr1.row; row <= cr2.row; row++) {
            for (col = cr1.col; col <= cr2.col; col++) {
               cr = SocialCalc.crToCoord(col, row);
               cell=sheet.GetAssuredCell(cr);
               if (cell.readonly) continue;
               if (saveundo) changes.AddUndo("set "+cr+" all", sheet.CellToString(cell));

               if (!sheet.cells[cr]) { // if had nothing
                  continue; // don't save anything
                  }
               movingcells[cr] = new SocialCalc.Cell(cr); // create new cell to copy

               for (attrib in cellProperties) { // go through each property
                  if (typeof cell[attrib] === undefined) { // don't copy undefined things and no need to delete
                     continue;
                     }
                  else {
                     movingcells[cr][attrib] = cell[attrib]; // copy for potential moving
                     }
                  if (rest == "all") {
                     delete cell[attrib];
                     }
                  if (rest == "formulas") {
                     if (cellProperties[attrib] == 1 || cellProperties[attrib] == 3) {
                        delete cell[attrib];
                        }
                     }
                  if (rest == "formats") {
                     if (cellProperties[attrib] == 2) {
                        delete cell[attrib];
                        }
                     }
                  }
               if (rest == "formulas") { // leave pristene deleted cell
                  cell.datavalue = "";
                  cell.datatype = null;
                  cell.formula = "";
                  cell.valuetype = "b";
                  }
               if (rest == "all") { // leave nothing for move all
                  delete sheet.cells[cr];
                  }
               }
            }

         // if moveinsert, check destination OK, and calculate pushing parameters

         if (cmd1 == "moveinsert") {
            inserthoriz = false;
            insertvert = false;
            if (rowoffset==0 && (destcr.col < cr1.col || destcr.col > cr2.col)) {
               if (destcr.col < cr1.col) { // moving left
                  pushamount = cr1.col - destcr.col;
                  inserthoriz = -1;
                  }
               else {
                  destcr.col -= 1;
                  coloffset = destcr.col - cr2.col;
                  pushamount = destcr.col - cr2.col;
                  inserthoriz = 1;
                  }
               }
            else if (coloffset==0 && (destcr.row < cr1.row || destcr.row > cr2.row)) {
               if (destcr.row < cr1.row) { // moving up
                  pushamount = cr1.row - destcr.row;
                  insertvert = -1;
                  }
               else {
                  destcr.row -= 1;
                  rowoffset = destcr.row - cr2.row;
                  pushamount = destcr.row - cr2.row;
                  insertvert = 1;
                  }
               }
            else {
               cmd1 = "movepaste"; // not allowed right now - ignore
               }                
            }

         // push any cells that need pushing

         movedto = {}; // remember what was moved where

         if (insertvert) {
            for (row = 0; row < pushamount; row++) {
               for (col = cr1.col; col <= cr2.col; col++) {
                  if (insertvert < 0) {
                     crbase = SocialCalc.crToCoord(col, destcr.row+pushamount-row-1); // from cell
                     cr = SocialCalc.crToCoord(col, cr2.row-row); // to cell
                     }
                  else {
                     crbase = SocialCalc.crToCoord(col, destcr.row-pushamount+row+1); // from cell
                     cr = SocialCalc.crToCoord(col, cr1.row+row); // to cell
                     }

                  basecell = sheet.GetAssuredCell(crbase);
                  if (saveundo) changes.AddUndo("set "+crbase+" all", sheet.CellToString(basecell));

                  cell = sheet.GetAssuredCell(cr);
                  if (rest == "all" || rest == "formats") {
                     for (attrib in cellProperties) {
                        if (cellProperties[attrib] == 1) continue; // copy only format attributes
                        if (typeof basecell[attrib] === undefined || cellProperties[attrib] == 3) {
                           delete cell[attrib];
                           }
                        else {
                           cell[attrib] = basecell[attrib];
                           }
                        }
                     }
                  if (rest == "all" || rest == "formulas") {
                     cell.datavalue = basecell.datavalue;
                     cell.datatype = basecell.datatype;            
                     cell.valuetype = basecell.valuetype;
                     cell.formula = basecell.formula;
                     delete cell.parseinfo;
                     cell.errors = basecell.errors;
                     }
                  delete cell.displaystring;

                  movedto[crbase] = cr; // old crbase is now at cr
                  }
               }
            }
         if (inserthoriz) {
            for (col = 0; col < pushamount; col++) {
               for (row = cr1.row; row <= cr2.row; row++) {
                  if (inserthoriz < 0) {
                     crbase = SocialCalc.crToCoord(destcr.col+pushamount-col-1, row);
                     cr = SocialCalc.crToCoord(cr2.col-col, row);
                     }
                  else {
                     crbase = SocialCalc.crToCoord(destcr.col-pushamount+col+1, row);
                     cr = SocialCalc.crToCoord(cr1.col+col, row);
                     }

                  basecell = sheet.GetAssuredCell(crbase);
                  if (saveundo) changes.AddUndo("set "+crbase+" all", sheet.CellToString(basecell));

                  cell = sheet.GetAssuredCell(cr);
                  if (rest == "all" || rest == "formats") {
                     for (attrib in cellProperties) {
                        if (cellProperties[attrib] == 1) continue; // copy only format attributes
                        if (typeof basecell[attrib] === undefined || cellProperties[attrib] == 3) {
                           delete cell[attrib];
                           }
                        else {
                           cell[attrib] = basecell[attrib];
                           }
                        }
                     }
                  if (rest == "all" || rest == "formulas") {
                     cell.datavalue = basecell.datavalue;
                     cell.datatype = basecell.datatype;            
                     cell.valuetype = basecell.valuetype;
                     cell.formula = basecell.formula;
                     delete cell.parseinfo;
                     cell.errors = basecell.errors;
                     }
                  delete cell.displaystring;

                  movedto[crbase] = cr; // old crbase is now at cr
                  }
               }
            }

         // paste moved cells into new place

         if (destcr.col+numcols-1 > attribs.lastcol) attribs.lastcol = destcr.col+numcols-1;
         if (destcr.row+numrows-1 > attribs.lastrow) attribs.lastrow = destcr.row+numrows-1;

         for (row = cr1.row; row < cr1.row+numrows; row++) {
            for (col = cr1.col; col < cr1.col+numcols; col++) {
               cr = SocialCalc.crToCoord(col+coloffset, row+rowoffset);
               cell=sheet.GetAssuredCell(cr);
               if (cell.readonly) continue;
               if (saveundo) changes.AddUndo("set "+cr+" all", sheet.CellToString(cell));

               crbase = SocialCalc.crToCoord(col, row); // get old cell to move

               movedto[crbase] = cr; // old crbase (moved cell) will now be at cr (destination)

               if (rest == "all" && !movingcells[crbase]) { // moving an empty cell
                  delete sheet.cells[cr]; // make the cell empty
                  continue;
                  }

               basecell = movingcells[crbase];
               if (!basecell) basecell = sheet.GetAssuredCell(crbase);

               if (rest == "all" || rest == "formats") {
                  for (attrib in cellProperties) {
                     if (cellProperties[attrib] == 1) continue; // copy only format attributes
                     if (typeof basecell[attrib] === undefined || cellProperties[attrib] == 3) {
                        delete cell[attrib];
                        }
                     else {
                        cell[attrib] = basecell[attrib];
                        }
                     }
                  }
               if (rest == "all" || rest == "formulas") {
                  cell.datavalue = basecell.datavalue;
                  cell.datatype = basecell.datatype;            
                  cell.valuetype = basecell.valuetype;
                  cell.formula = basecell.formula;
                  delete cell.parseinfo;
                  cell.errors = basecell.errors;
                  if (basecell.comment) { // comments are pasted as part of content, though not filled, etc.
                     cell.comment = basecell.comment;
                     }
                  else if (cell.comment) {
                     delete cell.comment;
                     }
                  }
               delete cell.displaystring;
               }
            }

         // do fixups

         for (cr in sheet.cells) { // update cell references to moved cells in calculated formulas
             cell = sheet.cells[cr];
             if (cell) {
                if (cell.datatype == "f") {
                   oldformula = cell.formula;
                   cell.formula = SocialCalc.ReplaceFormulaCoords(oldformula, movedto);
                   if (cell.formula != oldformula) {
                      delete cell.parseinfo;
                      if (saveundo && !movedto[cr]) { // moved cells are already saved for undo
                         changes.AddUndo("set "+cr+" formula "+oldformula);
                         }
                      }
                   }
                else {
                   delete cell.parseinfo;
                   }
                }
             }

         for (name in sheet.names) { // update cell references to moved cells in names
            if (sheet.names[name]) { // works with "A1", "A1:A20", and "=formula" forms
               v1 = sheet.names[name].definition;
               oldformula = v1;
               v2 = "";
               if (v1.charAt(0) == "=") {
                  v2 = "=";
                  v1 = v1.substring(1);
                  }
               sheet.names[name].definition = v2 +
                  SocialCalc.ReplaceFormulaCoords(v1, movedto);
               if (saveundo && sheet.names[name].definition != oldformula) { // save changes
                  changes.AddUndo("name define "+name+" "+oldformula);
                  }
               }
            }

         attribs.needsrecalc = "yes";
         break;

      case "name":
         what = cmd.NextToken();
         name = cmd.NextToken();
         rest = cmd.RestOfString();

         name = name.toUpperCase().replace(/[^A-Z0-9_\.]/g, "");
         if (name == "") break; // must have something

         if (what == "define") {
            if (rest == "") break; // must have something
            if (sheet.names[name]) { // already exists
               if (saveundo) changes.AddUndo("name define "+name+" "+sheet.names[name].definition);
               sheet.names[name].definition = rest;
               }
            else { // new
               if (saveundo) changes.AddUndo("name delete "+name);
               sheet.names[name] = {definition: rest, desc: ""};
               }
            }
         else if (what == "desc") {
            if (sheet.names[name]) { // must already exist
               if (saveundo) changes.AddUndo("name desc "+name+" "+sheet.names[name].desc);
               sheet.names[name].desc = rest;
               }
            }
         else if (what == "delete") {
            if (saveundo) {
               if (sheet.names[name].desc) changes.AddUndo("name desc "+name+" "+sheet.names[name].desc);
               changes.AddUndo("name define "+name+" "+sheet.names[name].definition);
               }
            delete sheet.names[name];
            }
         attribs.needsrecalc = "yes";

         break;

      case "recalc":
         attribs.needsrecalc = "yes"; // request recalc
         sheet.recalconce = true; // even if turned off
         break;

      case "redisplay":
         sheet.renderneeded = true;
         break;

      case "changedrendervalues": // needed for undo sometimes
         sheet.changedrendervalues = true;
         break;

      case "startcmdextension": // startcmdextension extension rest-of-command
         name = cmd.NextToken();
         cmdextension = sheet.sci.CmdExtensionCallbacks[name];
         if (cmdextension) {
            cmdextension.func(name, cmdextension.data, sheet, cmd, saveundo);
            }
         break;

      default:
         errortext = scc.s_escUnknownCmd+cmdstr;
         break;
      }

/* For Debugging:
var ustack="";
for (var i=0;i<sheet.changes.stack.length;i++) {
   ustack+=(i-0)+":"+sheet.changes.stack[i].command[0]+" of "+sheet.changes.stack[i].command.length+"/"+sheet.changes.stack[i].undo[0]+" of "+sheet.changes.stack[i].undo.length+",";
   }
alert(cmdstr+"|"+sheet.changes.stack.length+"--"+ustack);
*/

   return errortext;

   }

SocialCalc.SheetUndo = function(sheet) {

   var i;
   var tos = sheet.changes.TOS();
   var lastone = tos ? tos.undo.length-1 : -1;
   var cmdstr = "";

   for (i=lastone; i>=0; i--) { // do them backwards
      if (cmdstr) cmdstr += "\n"; // concatenate with separate lines
      cmdstr += tos.undo[i];
      }
   sheet.changes.Undo();
   sheet.ScheduleSheetCommands(cmdstr, false); // do undo operations

   }

SocialCalc.SheetRedo = function(sheet) {

   var tos, i;
   var didredo = sheet.changes.Redo();
   if (!didredo) {
      sheet.ScheduleSheetCommands("", false); // schedule doing nothing
      return;
      }
   tos = sheet.changes.TOS();
   var cmdstr = "";

   for (i=0; tos && i<tos.command.length; i++) {
      if (cmdstr) cmdstr += "\n"; // concatenate with separate lines
      cmdstr += tos.command[i];
      }
   sheet.ScheduleSheetCommands(cmdstr, false); // do undo operations

   }

SocialCalc.CreateAuditString = function(sheet) {

   var i, j;
   var result = "";
   var stack = sheet.changes.stack;
   var tos = sheet.changes.tos;
   for (i=0; i<=tos; i++) {
      for (j=0; j<stack[i].command.length; j++) {
         result += stack[i].command[j] + "\n";
         }
      }

   return result;

   }

SocialCalc.GetStyleNum = function(sheet, atype, style) {

   var num;

   if (style.length==0) return 0; // null means use zero, which means default or global default

   num = sheet[atype+"hash"][style];
   if (!num) {
      if (sheet[atype+"s"].length<1) sheet[atype+"s"].push("");
      num = sheet[atype+"s"].push(style) - 1;
      sheet[atype+"hash"][style] = num;
      sheet.changedrendervalues = true;
      }
   return num;

   }

SocialCalc.GetStyleString = function(sheet, atype, num) {

   if (!num) return null; // zero, null, and undefined return null

   return sheet[atype+"s"][num];

   }

//
// updatedformula = SocialCalc.OffsetFormulaCoords(formula, coloffset, rowoffset)
//
// Change relative cell references by offsets (even those to other worksheets so fill, paste, sort work as expected).
// If not what you want, use absolute references.
//

SocialCalc.OffsetFormulaCoords = function(formula, coloffset, rowoffset) {

   var parseinfo, ttext, ttype, i, cr, newcr;
   var updatedformula = "";
   var scf = SocialCalc.Formula;
   if (!scf) {
      return "Need SocialCalc.Formula";
      }
   var tokentype = scf.TokenType;
   var token_op = tokentype.op;
   var token_string = tokentype.string;
   var token_coord = tokentype.coord;
   var tokenOpExpansion = scf.TokenOpExpansion;

   parseinfo = scf.ParseFormulaIntoTokens(formula);

   for (i=0; i<parseinfo.length; i++) {
      ttype = parseinfo[i].type;
      ttext = parseinfo[i].text;
      if (ttype == token_coord) {
         newcr = "";
         cr = SocialCalc.coordToCr(ttext);
         if (ttext.charAt(0)!="$") { // add col offset unless absolute column
            cr.col += coloffset;
            }
         else {
            newcr += "$";
            }
         newcr += SocialCalc.rcColname(cr.col);
         if (ttext.indexOf("$", 1)==-1) { // add row offset unless absolute row
            cr.row += rowoffset;
            }
         else {
            newcr += "$";
            }
         newcr += cr.row;
         if (cr.row < 1 || cr.col < 1) {
            newcr = "#REF!";
            }
         updatedformula += newcr;
         }
      else if (ttype == token_string) {
         if (ttext.indexOf('"') >= 0) { // quotes to double
            updatedformula += '"' + ttext.replace(/"/, '""') + '"';
            }
         else updatedformula += '"' + ttext + '"';
         }
      else if (ttype == token_op) {
         updatedformula += tokenOpExpansion[ttext] || ttext; // make sure short tokens (e.g., "G") go back full (">=")
         }
      else { // leave everything else alone
         updatedformula += ttext;
         }
      }

   return updatedformula;

   }

//
// updatedformula = SocialCalc.AdjustFormulaCoords(formula, col, coloffset, row, rowoffset)
//
// Change all cell references to cells starting with col/row by offsets
//

SocialCalc.AdjustFormulaCoords = function(formula, col, coloffset, row, rowoffset) {

   var ttype, ttext, i, newcr;
   var updatedformula = "";
   var sheetref = false;
   var scf = SocialCalc.Formula;
   if (!scf) {
      return "Need SocialCalc.Formula";
      }
   var tokentype = scf.TokenType;
   var token_op = tokentype.op;
   var token_string = tokentype.string;
   var token_coord = tokentype.coord;
   var tokenOpExpansion = scf.TokenOpExpansion;

   parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(formula);

   for (i=0; i<parseinfo.length; i++) {
      ttype = parseinfo[i].type;
      ttext = parseinfo[i].text;
      if (ttype == token_op) { // references with sheet specifier are not offset
         if (ttext == "!") {
            sheetref = true; // found a sheet reference
            }
         else if (ttext != ":") { // for everything but a range, reset
            sheetref = false;
            }
         ttext = tokenOpExpansion[ttext] || ttext; // make sure short tokens (e.g., "G") go back full (">=")
         }
      if (ttype == token_coord) {
         cr = SocialCalc.coordToCr(ttext);
         if ((coloffset < 0 && cr.col >= col && cr.col < col-coloffset) ||
             (rowoffset < 0 && cr.row >= row && cr.row < row-rowoffset)) { // refs to deleted cells become invalid
            if (!sheetref) {
               cr.col = 0;
               cr.row = 0;
               }
            }
         if (!sheetref) {
            if (cr.col >= col) {
               cr.col += coloffset;
               }
            if (cr.row >= row) {
               cr.row += rowoffset;
               }
            }
         if (ttext.charAt(0)=="$") {
            newcr = "$"+SocialCalc.rcColname(cr.col);
            }
         else {
            newcr = SocialCalc.rcColname(cr.col);
            }
         if (ttext.indexOf("$", 1)!=-1) {
            newcr += "$" + cr.row;
            }
         else {
            newcr += cr.row;
            }
         if (cr.row < 1 || cr.col < 1) {
            newcr = "#REF!";
            }
         ttext = newcr;
         }
      else if (ttype == token_string) {
         if (ttext.indexOf('"') >= 0) { // quotes to double
            ttext = '"' + ttext.replace(/"/, '""') + '"';
            }
         else ttext = '"' + ttext + '"';
         }
      updatedformula += ttext;
      }

   return updatedformula;

   }

//
// updatedformula = SocialCalc.ReplaceFormulaCoords(formula, movedto)
//
// Change all cell references to cells that are keys in moveto to be to moveto[coord].
// Don't change references to other sheets.
// Handle range extents specially.
//

SocialCalc.ReplaceFormulaCoords = function(formula, movedto) {

   var ttype, ttext, i, newcr, coord;
   var updatedformula = "";
   var sheetref = false;
   var scf = SocialCalc.Formula;
   if (!scf) {
      return "Need SocialCalc.Formula";
      }
   var tokentype = scf.TokenType;
   var token_op = tokentype.op;
   var token_string = tokentype.string;
   var token_coord = tokentype.coord;
   var tokenOpExpansion = scf.TokenOpExpansion;

   parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(formula);

   for (i=0; i<parseinfo.length; i++) {
      ttype = parseinfo[i].type;
      ttext = parseinfo[i].text;
      if (ttype == token_op) { // references with sheet specifier are not change
         if (ttext == "!") {
            sheetref = true; // found a sheet reference
            }
         else if (ttext != ":") { // for everything but a range, reset
            sheetref = false;
            }

//!!!! HANDLE RANGE EXTENT MOVES

         ttext = tokenOpExpansion[ttext] || ttext; // make sure short tokens (e.g., "G") go back full (">=")
         }
      if (ttype == token_coord) {
         cr = SocialCalc.coordToCr(ttext); // get parts
         coord = SocialCalc.crToCoord(cr.col, cr.row); // get "clean" reference
         if (movedto[coord] && !sheetref) { // this is a reference to a moved cell
            cr = SocialCalc.coordToCr(movedto[coord]); // get new row and col
            if (ttext.charAt(0)=="$") { // copy absolute ref marks if present
               newcr = "$"+SocialCalc.rcColname(cr.col);
               }
            else {
               newcr = SocialCalc.rcColname(cr.col);
               }
            if (ttext.indexOf("$", 1)!=-1) {
               newcr += "$" + cr.row;
               }
            else {
               newcr += cr.row;
               }
            ttext = newcr;
            }
         }
      else if (ttype == token_string) {
         if (ttext.indexOf('"') >= 0) { // quotes to double
            ttext = '"' + ttext.replace(/"/, '""') + '"';
            }
         else ttext = '"' + ttext + '"';
         }
      updatedformula += ttext;
      }

   return updatedformula;

   }


// ************************
//
// Recalc Loop Code
//
// ************************

//
// How recalc works:
//
// !!!!!!!!!!!!!!
//

// SocialCalc.RecalcInfo - object with global recalc info

SocialCalc.RecalcInfo = {

   sheet: null, // which sheet is being recalced

   currentState: 0, // current state
   state: {idle: 0, start_calc: 1, order: 2, calc: 3, start_wait: 4, done_wait: 5}, // allowed state values

   recalctimer: null, // value to cancel timer
   maxtimeslice: 100, // maximum milliseconds per slice of recalc time before a wait
   timeslicedelay: 1, // milliseconds to wait between recalc time slices
   starttime: 0, // when recalc started

   queue: [], // queue of sheet waiting to be recalced

   // LoadSheet: a function that returns true if started a load or false if not.
   //

   LoadSheet: function(sheetname) {return false;} // default returns not found

   }

// SocialCalc.RecalcData - object with recalc info while determining recalc order and afterward

SocialCalc.RecalcData = function() { // initialize a RecalcData object

   this.inrecalc = true; // if true, doing a recalc
   this.celllist = []; // list with all potential cells to calculate
   this.celllistitem = 0; // cell to check next when ordering
   this.calclist = null; // object which is the chained list of cells to calculate
                         // each in the form of "coord: nextcoord"
                         // e.g., if B8 is calculated right after A8, then calclist.A8=="B8"
                         // if null, need to create the list
   this.calclistlength = 0; // number of items in calclist

   this.firstcalc = null; // start of the calc list - a string or null
   this.lastcalc = null; // last one on chain (used to add more to the end)

   this.nextcalc = null; // used to keep track during background recalc to make it restartable
   this.count = 0; // number calculated

   // checkinfo is used when determining calc order:

   this.checkinfo = {}; // attributes are coords; if no attrib for a coord, it wasn't checked or doesn't need it
                        // values are RecalcCheckInfo objects while checking or TRUE when complete

   }

// SocialCalc.RecalcCheckInfo - object that stores checking info while determining recalc order

SocialCalc.RecalcCheckInfo = function() { // initialize a RecalcCheckInfo object

   this.oldcoord = null; // chain back up of cells referring to cells
   this.parsepos = 0; // which token we are up to

   // range info

   this.inrange = false; // if true, in the process of checking a range of coords
   this.inrangestart = false; // if true, have not yet filled in range loop values
   this.cr1 = null; // range first coord as a cr object
   this.cr2 = null; // range second coord as a cr object
   this.c1 = null; // range extents
   this.c2 = null;
   this.r1 = null;
   this.r2 = null;
   this.c = null; // looping values
   this.r = null;
   
   }

// Recalc the entire sheet

SocialCalc.RecalcSheet = function(sheet) {

   var coord, err, recalcdata;
   var scri = SocialCalc.RecalcInfo;

   if (scri.currentState != scri.state.idle) {
      scri.queue.push(sheet);
      return;
      }

   delete sheet.attribs.circularreferencecell; // reset recalc-wide things
   SocialCalc.Formula.FreshnessInfoReset();

   SocialCalc.RecalcClearTimeout();

   scri.sheet = sheet; // set values needed by background recalc
   scri.currentState = scri.state.start_calc;

   scri.starttime = new Date();

   if (sheet.statuscallback) {
      sheet.statuscallback(scri, "calcstart", null, sheet.statuscallbackparams);
      }

   SocialCalc.RecalcSetTimeout();

   }

//
// SocialCalc.RecalcSetTimeout - set a timer for next recalc step
//

SocialCalc.RecalcSetTimeout = function() {

   var scri = SocialCalc.RecalcInfo;

   scri.recalctimer = window.setTimeout(SocialCalc.RecalcTimerRoutine, scri.timeslicedelay);

   }

//
// SocialCalc.RecalcClearTimeout - cancel any timeouts
//

SocialCalc.RecalcClearTimeout = function() {

   var scri = SocialCalc.RecalcInfo;

   if (scri.recalctimer) {
      window.clearTimeout(scri.recalctimer);
      scri.recalctimer = null;
      }

   }


//
// SocialCalc.RecalcLoadedSheet(sheetname, str, recalcneeded, live)
//
// Called when a sheet finishes loading with name, string, and t/f whether it should be recalced.
// If loaded sheet has sheet.attribs.recalc=="off", then no recalc done.
// If sheetname is null, then the sheetname waiting for will be used.
//

SocialCalc.RecalcLoadedSheet = function(sheetname, str, recalcneeded, live) {

   var sheet;
   var scri = SocialCalc.RecalcInfo;
   var scf = SocialCalc.Formula;

   sheet = SocialCalc.Formula.AddSheetToCache(sheetname || scf.SheetCache.waitingForLoading, str, live);

   if (recalcneeded && sheet && sheet.attribs.recalc!="off") { // if recalcneeded, and not manual sheet, chain in this new sheet to recalc loop
      sheet.previousrecalcsheet = scri.sheet;
      scri.sheet = sheet;
      scri.currentState = scri.state.start_calc;
      }
   scf.SheetCache.waitingForLoading = null;

   SocialCalc.RecalcSetTimeout();

   }


//
// SocialCalc.RecalcTimerRoutine - handles the actual order determination and cell-by-cell recalculation in the background
//

SocialCalc.RecalcTimerRoutine = function() {

   var eresult, cell, coord, err, status;
   var starttime = new Date();
   var count = 0;
   var scf = SocialCalc.Formula;
   if (!scf) {
      return "Need SocialCalc.Formula";
      }
   var scri = SocialCalc.RecalcInfo;
   var sheet = scri.sheet;
   if (!sheet) {
      return;
      }
   var recalcdata = sheet.recalcdata || (sheet.recalcdata = {});

   var do_statuscallback = function(status, arg) { // routine to do callback if required
      if (sheet.statuscallback) {
         sheet.statuscallback(recalcdata, status, arg, sheet.statuscallbackparams);
         }
      }

   SocialCalc.RecalcClearTimeout();

   if (scri.currentState == scri.state.start_calc) {

      recalcdata = new SocialCalc.RecalcData();
      sheet.recalcdata = recalcdata;

      for (coord in sheet.cells) { // get list of cells to check for order
         if (!coord) continue;
         recalcdata.celllist.push(coord);
         }

      recalcdata.calclist = {}; // start with empty list
      scri.currentState = scri.state.order; // drop through to determining recalc order
      }

   if (scri.currentState == scri.state.order) {
      while (recalcdata.celllistitem < recalcdata.celllist.length) { // check all the cells to see if they should be on the list
         coord = recalcdata.celllist[recalcdata.celllistitem++];
         err = SocialCalc.RecalcCheckCell(sheet, coord);
         if (((new Date()) - starttime) >= scri.maxtimeslice) { // if taking too long, give up CPU for a while
            do_statuscallback("calcorder", {coord: coord, total: recalcdata.celllist.length, count: recalcdata.celllistitem});
            SocialCalc.RecalcSetTimeout();
            return;
            }
         }

      do_statuscallback("calccheckdone", recalcdata.calclistlength);

      recalcdata.nextcalc = recalcdata.firstcalc; // start at the beginning of the recalc chain
      scri.currentState = scri.state.calc; // loop through cells on next timer call
      SocialCalc.RecalcSetTimeout();
      return;
      }

   if (scri.currentState == scri.state.start_wait) { // starting to wait for something
      scri.currentState = scri.state.done_wait; // finished on next timer call
      if (scri.LoadSheet) {
         status = scri.LoadSheet(scf.SheetCache.waitingForLoading);
         if (status) { // started a load operation
            return;
            }
         }
      SocialCalc.RecalcLoadedSheet(null, "", false);
      return;
      }

   if (scri.currentState == scri.state.done_wait) {
      scri.currentState = scri.state.calc; // loop through cells on next timer call
      SocialCalc.RecalcSetTimeout();
      return;
      }

   // otherwise should be scri.state.calc

   if (scri.currentState != scri.state.calc) {
      alert("Recalc state error: "+scri.currentState+". Error in SocialCalc code.");
      }

   coord = sheet.recalcdata.nextcalc;
   while (coord) {
      cell = sheet.cells[coord];
      eresult = scf.evaluate_parsed_formula(cell.parseinfo, sheet, false);
      if (scf.SheetCache.waitingForLoading) { // wait until restarted
         recalcdata.nextcalc = coord; // start with this cell again
         recalcdata.count += count;
         do_statuscallback("calcloading", {sheetname: scf.SheetCache.waitingForLoading});
         scri.currentState = scri.state.start_wait; // start load on next timer call
         SocialCalc.RecalcSetTimeout();
         return;
         }

      if (scf.RemoteFunctionInfo.waitingForServer) { // wait until restarted
         recalcdata.nextcalc = coord; // start with this cell again
         recalcdata.count += count;
         do_statuscallback("calcserverfunc",
            {funcname: scf.RemoteFunctionInfo.waitingForServer, coord: coord, total: recalcdata.calclistlength, count: recalcdata.count});
         scri.currentState = scri.state.done_wait; // start load on next timer call
         return; // return and wait for next recalc timer event
         }

      if (cell.datavalue != eresult.value ||
       cell.valuetype != eresult.type) { // only update if changed from last time
         cell.datavalue = eresult.value;
         cell.valuetype = eresult.type;
         delete cell.displaystring;
         sheet.recalcchangedavalue = true; // remember something changed in case other code wants to know
         }
      if (eresult.error) {
         cell.errors = eresult.error;
         }
      count++;
      coord = sheet.recalcdata.calclist[coord];

      if (((new Date()) - starttime) >= scri.maxtimeslice) { // if taking too long, give up CPU for a while
         recalcdata.nextcalc = coord; // start with next cell on chain
         recalcdata.count += count;
         do_statuscallback("calcstep", {coord: coord, total: recalcdata.calclistlength, count: recalcdata.count});
         SocialCalc.RecalcSetTimeout();
         return;
         }
      }

   recalcdata.inrecalc = false;

   delete sheet.recalcdata; // save memory and clear out for name lookup formula evaluation

   delete sheet.attribs.needsrecalc; // remember recalc done

   scri.sheet = sheet.previousrecalcsheet || null; // chain back if doing recalc of loaded sheets
   if (scri.sheet) {
      scri.currentState = scri.state.calc; // start where we left off
      SocialCalc.RecalcSetTimeout();
      return;
      }

   scf.FreshnessInfo.recalc_completed = true; // say freshness info is complete
   scri.currentState = scri.state.idle; // we are idle

   do_statuscallback("calcfinished", (new Date()) - scri.starttime);

   // Check queue for more sheets.
   if (scri.queue.length > 0) {
      sheet = scri.queue.shift();
      sheet.RecalcSheet();
      }
   }


//
// circref = SocialCalc.RecalcCheckCell(sheet, coord)
//
// Checks cell to put on calclist, looking at parsed tokens.
// Also checks cells this cell is dependent upon
// if it contains a formula with cell references.
// If circular reference, returns non-null.
//

SocialCalc.RecalcCheckCell = function(sheet, startcoord) {

   var parseinfo, ttext, ttype, i, rangecoord, circref, value, pos, pos2, cell, coordvals;
   var scf = SocialCalc.Formula;
   if (!scf) {
      return "Need SocialCalc.Formula";
      }
   var tokentype = scf.TokenType;
   var token_op = tokentype.op;
   var token_name = tokentype.name;
   var token_coord = tokentype.coord;

   var recalcdata = sheet.recalcdata;
   var checkinfo = recalcdata.checkinfo;

   var sheetref = false; // if true, a sheet reference is in effect, so don't check that
   var oldcoord = null; // coord of formula that referred to this one when checking down the tree
   var coord = startcoord; // the coord of the cell we are checking

   // Start with requested cell, and then continue down or up the dependency tree
   // oldcoord (and checkinfo[coord].oldcoord) maintains the reference stack during the tree walk
   // checkinfo[coord] maintains the stack of checking looping values, e.g., token number being checked

mainloop:
   while (coord) {
      cell = sheet.cells[coord];
      coordvals = checkinfo[coord];

      if (!cell || cell.datatype != "f" || // Don't calculate if not a formula
          (coordvals && typeof coordvals != "object")) { // Don't calc if already calculated
         coord = oldcoord; // go back up dependency tree to coord that referred to us
         if (checkinfo[coord]) oldcoord = checkinfo[coord].oldcoord;
         continue;
         }

      if (!coordvals) { // do we have checking information about this cell?
         coordvals = new SocialCalc.RecalcCheckInfo(); // no - make a place to hold it
         checkinfo[coord] = coordvals;
         }

      if (cell.errors) { // delete errors from previous recalcs
         delete cell.errors;
         }

      if (!cell.parseinfo) { // cache parsed formula
         cell.parseinfo = scf.ParseFormulaIntoTokens(cell.formula);
         }
      parseinfo = cell.parseinfo;

      for (i=coordvals.parsepos; i<parseinfo.length; i++) { // go through each token in formula

         if (coordvals.inrange) { // processing a range of coords
            if (coordvals.inrangestart) { // first time - fill in other values
               if (coordvals.cr1.col > coordvals.cr2.col) { coordvals.c1 = coordvals.cr2.col; coordvals.c2 = coordvals.cr1.col; }
               else { coordvals.c1 = coordvals.cr1.col; coordvals.c2 = coordvals.cr2.col; }
               coordvals.c = coordvals.c1 - 1; // start one before

               if (coordvals.cr1.row > coordvals.cr2.row) { coordvals.r1 = coordvals.cr2.row; coordvals.r2 = coordvals.cr1.row; }
               else { coordvals.r1 = coordvals.cr1.row; coordvals.r2 = coordvals.cr2.row; }
               coordvals.r = coordvals.r1; // start on this row
               coordvals.inrangestart = false;
               }
            else { // not first time
               }
            coordvals.c += 1; // increment column
            if (coordvals.c > coordvals.c2) { // finished the columns of this row
               coordvals.r += 1; // increment row
               if (coordvals.r > coordvals.r2) { // finished checking the entire range
                  coordvals.inrange = false;
                  continue;
                  }
               coordvals.c = coordvals.c1; // start at the beginning of next row
               }
            rangecoord = SocialCalc.crToCoord(coordvals.c, coordvals.r);

            // now check that one

            coordvals.parsepos = i; // remember our position
            coordvals.oldcoord = oldcoord; // remember back up chain
            oldcoord = coord; // come back to us
            coord = rangecoord;
            if (checkinfo[coord] && typeof checkinfo[coord] == "object") { // Circular reference
               cell.errors = SocialCalc.Constants.s_caccCircRef+startcoord; // set on original cell making the ref
               checkinfo[startcoord] = true; // this one should be calculated once at this point
               if (!recalcdata.firstcalc) {
                  recalcdata.firstcalc = startcoord;
                  }
               else {
                  recalcdata.calclist[recalcdata.lastcalc] = startcoord;
                  }
               recalcdata.lastcalc = startcoord;
               recalcdata.calclistlength++; // count number on list
               sheet.attribs.circularreferencecell = coord+"|"+oldcoord; // remember at least one circ ref
               return cell.errors;
               }
            continue mainloop;
            }

         ttype = parseinfo[i].type; // get token details
         ttext = parseinfo[i].text;
         if (ttype == token_op) { // references with sheet specifier are not checked
            if (ttext == "!") {
               sheetref = true; // found a sheet reference
               }
            else if (ttext != ":") { // for everything but a range, reset
               sheetref = false;
               }
            }

         if (ttype == token_name) { // look for named range
            value = scf.LookupName(sheet, ttext);
            if (value.type == "range") { // only need to recurse here for range, which may be just one cell
               pos = value.value.indexOf("|");
               if (pos != -1) { // range - check each cell
                  coordvals.cr1 = SocialCalc.coordToCr(value.value.substring(0,pos));
                  pos2 = value.value.indexOf("|", pos+1);
                  coordvals.cr2 = SocialCalc.coordToCr(value.value.substring(pos+1,pos2));
                  coordvals.inrange = true;
                  coordvals.inrangestart = true;
                  i = i-1; // back up so will start up again here
                  continue;
                  }
               }
            else if (value.type == "coord") { // just a coord
               ttype = token_coord; // treat as a coord inline
               ttext = value.value; // and then drop through to next test which should succeed
               }
            else { // not a defined name - probably a function
               }
            }

         if (ttype == token_coord) { // token is a coord

            if (i >= 2 // look for a range
             && parseinfo[i-1].type == token_op && parseinfo[i-1].text == ':'
             && parseinfo[i-2].type == token_coord
             && !sheetref) { // Range -- check each cell
               coordvals.cr1 = SocialCalc.coordToCr(parseinfo[i-2].text); // remember range extents
               coordvals.cr2 = SocialCalc.coordToCr(ttext);
               coordvals.inrange = true; // next time use the range looping code
               coordvals.inrangestart = true;
               i = i-1; // back up so will start up again here
               continue;
               }

            else if (!sheetref) { // Single cell reference
               if (ttext.indexOf("$") != -1) ttext = ttext.replace(/\$/g, ""); // remove any $'s
               coordvals.parsepos = i+1; // remember our position - come back on next token
               coordvals.oldcoord = oldcoord; // remember back up chain
               oldcoord = coord; // come back to us
               coord = ttext;
               if (checkinfo[coord] && typeof checkinfo[coord] == "object") { // Circular reference
                  cell.errors = SocialCalc.Constants.s_caccCircRef+startcoord; // set on original cell making the ref
                  checkinfo[startcoord] = true; // this one should be calculated once at this point
                  if (!recalcdata.firstcalc) { // add to calclist
                     recalcdata.firstcalc = startcoord;
                     }
                  else {
                     recalcdata.calclist[recalcdata.lastcalc] = startcoord;
                     }
                  recalcdata.lastcalc = startcoord;
                  recalcdata.calclistlength++; // count number on list
                  sheet.attribs.circularreferencecell = coord+"|"+oldcoord; // remember at least one circ ref
                  return cell.errors;
                  }
               continue mainloop;
               }
            }
         }

      sheetref = false; // make sure off when bump back up

      checkinfo[coord] = true; // this one is finished
      if (!recalcdata.firstcalc) { // add to calclist
         recalcdata.firstcalc = coord;
         }
      else {
         recalcdata.calclist[recalcdata.lastcalc] = coord;
         }
      recalcdata.lastcalc = coord;
      recalcdata.calclistlength++; // count number on list

      coord = oldcoord; // go back to the formula that referred to us and continue
      oldcoord = checkinfo[coord] ? checkinfo[coord].oldcoord : null;

      }

   return "";

   }


// *************************************
//
// Parse class:
//
// Used by ExecuteSheetCommand to get elements of commands to execute.
// The string it works with consists of one or more lines each
// made up of one or more tokens separated by a delimiter.
//
// *************************************

// Initialize: set string to work with

SocialCalc.Parse = function(str) {

   // properties:

   this.str = str;
   this.pos = 0;
   this.delimiter = " ";
   this.lineEnd = str.indexOf("\n");
   if (this.lineEnd < 0) {
      this.lineEnd = str.length;
      }

   }

// Return next token as a string

SocialCalc.Parse.prototype.NextToken = function() {
   if (this.pos < 0) return "";
   var pos2 = this.str.indexOf(this.delimiter, this.pos);
   var pos1 = this.pos;
   if (pos2 > this.lineEnd) { // don't go past end of line
      pos2 = this.lineEnd;
      }
   if (pos2 >= 0) {
      this.pos = pos2 + 1;
      return this.str.substring(pos1, pos2);
      }
   else {
      this.pos = this.lineEnd;
      return this.str.substring(pos1, this.lineEnd);
      }
   }

// Return everything from current point until end of line

SocialCalc.Parse.prototype.RestOfString = function() {
   var oldpos = this.pos;
   if (this.pos < 0 || this.pos >= this.lineEnd) return "";
   this.pos = this.lineEnd;
   return this.str.substring(oldpos, this.lineEnd);
   }

SocialCalc.Parse.prototype.RestOfStringNoMove = function() {
   if (this.pos < 0 || this.pos >= this.lineEnd) return "";
   return this.str.substring(this.pos, this.lineEnd);
   }

// Move current point to next line

SocialCalc.Parse.prototype.NextLine = function() {
   this.pos = this.lineEnd + 1;
   this.lineEnd = this.str.indexOf("\n", this.pos);
   if (this.lineEnd < 0) {
      this.lineEnd = this.str.length;
      }
   }

// Check to see if at end of string with no more to process

SocialCalc.Parse.prototype.EOF = function() {
   if (this.pos < 0 || this.pos >= this.str.length) return true;
   return false;
   }


// *************************************
//
// UndoStack class:
//
// Implements the behavior needed for a normal application's undo/redo stack.
// You add a new change sequence with PushChange.
// The type argument is a string that can be used to lookup some general string 
// like "typing" or "setting attribute" for the menu prompts for undo/redo.
//
// You add the "do" steps with AddDo. The non-null, non-undefined arguments are
// joined together with " " to make a command string to be saved.
//
// You add the undo steps as commands for the most recent change with AddUndo.
// The non-null, non-undefined arguments are joined together with " " to make
// a command string to be saved.
//
// The Undo and Redo functions move the Top Of Stack pointer through the changes stack
// so you can undo and redo. Doing a new PushChange removes all undone items
// after TOS.
//
// You can push more things than you can undo if you want.
// There is a maximum to remember as the "did" stack for an audit trail (and as redo). This may be unlimited.
// There is a separate maximum to remember that can be undone. This may be smaller than maxRedo.
//
// *************************************

SocialCalc.UndoStack = function() {

   // properties:

   this.stack = []; // {command: [], type: type, undo: []} -- multiple dos and undos allowed
   this.tos = -1; // top of stack position, used for undo/redo
   this.maxRedo = 0; // Maximum size of redo stack (and audit trail which is this.stack[n].command) or zero if no limit
   this.maxUndo = 50; // Maximum number of steps kept for undo (usually the memory intensive part) or zero if no limit

   }

SocialCalc.UndoStack.prototype.PushChange = function(type) { // adding a new thing to the stack
   while (this.stack.length > 0 && this.stack.length-1 > this.tos) { // pop off things not redone
      this.stack.pop();
      }
   this.stack.push({command: [], type: type, undo: []});
   if (this.maxRedo && this.stack.length > this.maxRedo) { // limit number kept as audit trail
      this.stack.shift(); // remove the extra one
      }
   if (this.maxUndo && this.stack.length > this.maxUndo) { // need to trim excess undo info
      this.stack[this.stack.length - this.maxUndo - 1].undo = []; // only need to remove one
      }
   this.tos = this.stack.length - 1;
   }

SocialCalc.UndoStack.prototype.AddDo = function() {
   if (!this.stack[this.stack.length-1]) { return; }
   var args = [];
   for (var i=0; i<arguments.length; i++) {
      if (arguments[i]!=null) args.push(arguments[i]); // ignore null or undefined
      }
   var cmd = args.join(" ");
   this.stack[this.stack.length-1].command.push(cmd);
   }

SocialCalc.UndoStack.prototype.AddUndo = function() {
   if (!this.stack[this.stack.length-1]) { return; }
   var args = [];
   for (var i=0; i<arguments.length; i++) {
      if (arguments[i]!=null) args.push(arguments[i]); // ignore null or undefined
      }
   var cmd = args.join(" ");
   this.stack[this.stack.length-1].undo.push(cmd);
   }

SocialCalc.UndoStack.prototype.TOS = function() {
   if (this.tos >= 0) return this.stack[this.tos];
   else return null;
   }

SocialCalc.UndoStack.prototype.Undo = function() {
   if (this.tos >= 0 && (!this.maxUndo || this.tos > this.stack.length - this.maxUndo - 1)) {
      this.tos -= 1;
      return true;
      }
   else {
      return false;
      }
   }

SocialCalc.UndoStack.prototype.Redo = function() {
   if (this.tos < this.stack.length-1) {
      this.tos += 1;
      return true;
      }
   else {
      return false;
      }
   }

// *************************************
//
// Clipboard Object:
//
// This is a single object.
// Stores the clipboard, which is shared by all active sheets.
// Like the undo stack, it does not persist from one editing session to another.
//
// *************************************

SocialCalc.Clipboard = {

   // properties:

   clipboard:  "" // empty or string in save format with "copiedfrom:" set to a range

   }

// Functions:

SocialCalc.PrecomputeSheetFontsAndLayouts = function(context) {

   var defaultfont, parts, layoutre, dparts, sparts, num, s, i;
   var sheetobj = context.sheetobj;
   var attribs =  sheetobj.attribs;

   if (attribs.defaultfont) {
      defaultfont = sheetobj.fonts[attribs.defaultfont];
      defaultfont = defaultfont.replace(/^\*/,SocialCalc.Constants.defaultCellFontStyle);
      defaultfont = defaultfont.replace(/(.+)\*(.+)/,"$1"+SocialCalc.Constants.defaultCellFontSize+"$2");
      defaultfont = defaultfont.replace(/\*$/,SocialCalc.Constants.defaultCellFontFamily);
      parts=defaultfont.match(/^(\S+? \S+?) (\S+?) (\S.*)$/);
      context.defaultfontstyle = parts[1];
      context.defaultfontsize = parts[2];
      context.defaultfontfamily = parts[3];
   }

   for (num=1; num<sheetobj.fonts.length; num++) { // precompute fonts by filling in the *'s
      s=sheetobj.fonts[num];
      s=s.replace(/^\*/,context.defaultfontstyle);
      s=s.replace(/(.+)\*(.+)/,"$1"+context.defaultfontsize+"$2");
      s=s.replace(/\*$/,context.defaultfontfamily);
      parts=s.match(/^(\S+?) (\S+?) (\S+?) (\S.*)$/);
      context.fonts[num] = {style: parts[1], weight: parts[2], size: parts[3], family: parts[4]};

   }

   layoutre = /^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);/;
   dparts = SocialCalc.Constants.defaultCellLayout.match(layoutre); // get built-in defaults

   if (attribs.defaultlayout) {
      sparts = sheetobj.layouts[attribs.defaultlayout].match(layoutre); // get sheet defaults, if set
   }
   else {
      sparts = ["", "*", "*", "*", "*", "*"];
   }

   for (num=1; num<sheetobj.layouts.length; num++) { // precompute layouts by filling in the *'s
      s=sheetobj.layouts[num];
      parts = s.match(layoutre);
      for (i=1; i<=5; i++) {
         if (parts[i]=="*") {
            parts[i] = (sparts[i] != "*" ? sparts[i] : dparts[i]); // if *, sheet default or built-in
         }
      }
      context.layouts[num] = "padding:"+parts[1]+" "+parts[2]+" "+parts[3]+" "+parts[4]+
          ";vertical-align:"+parts[5]+";";
   }

   context.needprecompute = false;

}

SocialCalc.CalculateCellSkipData = function(context) {

   var row, col, coord, cell, contextcell, colspan, rowspan, skiprow, skipcol, skipcoord;

   var sheetobj=context.sheetobj;
   var sheetrowattribs=sheetobj.rowattribs;
   var sheetcolattribs=sheetobj.colattribs;
   context.maxrow=0;
   context.maxcol=0;
   context.cellskip = {}; // reset

   // Calculate cellskip data

   for (row=1; row<=sheetobj.attribs.lastrow; row++) {
      for (col=1; col<=sheetobj.attribs.lastcol; col++) { // look for spans and set cellskip for skipped cells
         coord=SocialCalc.crToCoord(col, row);
         cell=sheetobj.cells[coord];
         // don't look at undefined cells (they have no spans) or skipped cells
         if (cell===undefined || context.cellskip[coord]) continue;
         colspan=cell.colspan || 1;
         rowspan=cell.rowspan || 1;
         if (colspan>1 || rowspan>1) {
            for (skiprow=row; skiprow<row+rowspan; skiprow++) {
               for (skipcol=col; skipcol<col+colspan; skipcol++) { // do the setting on individual cells
                  skipcoord=SocialCalc.crToCoord(skipcol,skiprow);
                  if (skipcoord==coord) { // for coord, remember row and col
                     context.coordToCR[coord]={row: row, col: col};
                  }
                  else { // for other cells, flag with coord of here
                     context.cellskip[skipcoord]=coord;
                  }
                  if (skiprow>context.maxrow) maxrow=skiprow;
                  if (skipcol>context.maxcol) maxcol=skipcol;
               }
            }
         }
      }
   }

   context.needcellskip = false;

}

SocialCalc.CalculateColWidthData = function(context) {

   var colnum, colname, colwidth, totalwidth;

   var sheetobj=context.sheetobj;
   var sheetcolattribs=sheetobj.colattribs;

   // Calculate column width data

   totalwidth=context.showRCHeaders ? context.rownamewidth-0 : 0;
   for (colpane=0; colpane<context.colpanes.length; colpane++) {
      for (colnum=context.colpanes[colpane].first; colnum<=context.colpanes[colpane].last; colnum++) {
         colname=SocialCalc.rcColname(colnum);
         if (sheetobj.colattribs.hide[colname] == "yes") {
            context.colwidth[colnum] = 0;
         }
         else {
            colwidth = sheetobj.colattribs.width[colname] || sheetobj.attribs.defaultcolwidth || SocialCalc.Constants.defaultColWidth;
            if (colwidth=="blank" || colwidth=="auto") colwidth="";
            context.colwidth[colnum]=colwidth+"";
            totalwidth+=(colwidth && ((colwidth-0)>0)) ? (colwidth-0) : 10;
         }
      }
   }
   context.totalwidth = totalwidth;

}


// *************************************
//
// Misc. functions:
//
// *************************************

SocialCalc.rcColname = function(c) {
   if (c > 702) c = 702; // maximum number of columns - ZZ
   if (c < 1) c = 1;
   var collow = (c - 1) % 26 + 65;
   var colhigh = Math.floor((c - 1) / 26);
   if (colhigh)
      return String.fromCharCode(colhigh + 64) + String.fromCharCode(collow);
   else
      return String.fromCharCode(collow);
}

SocialCalc.letters = ["A","B","C","D","E","F","G","H","I","J","K","L","M",
   "N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];

SocialCalc.crToCoord = function(c, r) {
   var result;
   if (c < 1) c = 1;
   if (c > 702) c = 702; // maximum number of columns - ZZ
   if (r < 1) r = 1;
   var collow = (c - 1) % 26;
   var colhigh = Math.floor((c - 1) / 26);
   if (colhigh)
      result = SocialCalc.letters[colhigh-1] + SocialCalc.letters[collow] + r;
   else
      result = SocialCalc.letters[collow] + r;
   return result;
}

SocialCalc.coordToCol = {}; // too expensive to set in crToCoord since that is called so many times
SocialCalc.coordToRow = {};

SocialCalc.coordToCr = function(cr) {
   var c, i, ch;
   var r = SocialCalc.coordToRow[cr];
   if (r) return {row: r, col: SocialCalc.coordToCol[cr]};
   c=0;r=0;
   for (i=0; i<cr.length; i++) { // this was faster than using regexes; assumes well-formed
      ch = cr.charCodeAt(i);
      if (ch==36) ; // skip $'s
      else if (ch<=57) r = 10*r + ch-48;
      else if (ch>=97) c = 26*c + ch-96;
      else if (ch>=65) c = 26*c + ch-64;
   }
   SocialCalc.coordToCol[cr] = c;
   SocialCalc.coordToRow[cr] = r;
   return {row: r, col: c};

}

SocialCalc.ParseRange = function(range) {
   var pos, cr, cr1, cr2;
   if (!range) range = "A1:A1"; // error return, hopefully benign
   range = range.toUpperCase();
   pos = range.indexOf(":");
   if (pos>=0) {
      cr = range.substring(0,pos);
      cr1 = SocialCalc.coordToCr(cr);
      cr1.coord = cr;
      cr = range.substring(pos+1);
      cr2 = SocialCalc.coordToCr(cr);
      cr2.coord = cr;
   }
   else {
      cr1 = SocialCalc.coordToCr(range);
      cr1.coord = range;
      cr2 = SocialCalc.coordToCr(range);
      cr2.coord = range;
   }
   return {cr1: cr1, cr2: cr2};
}

SocialCalc.decodeFromSave = function(s) {
   if (typeof s != "string") return s;
   if (s.indexOf("\\")==-1) return s; // for performace reasons: replace nothing takes up time
   var r=s.replace(/\\c/g,":");
   r=r.replace(/\\n/g,"\n");
   return r.replace(/\\b/g,"\\");
}

SocialCalc.decodeFromAjax = function(s) {
   if (typeof s != "string") return s;
   if (s.indexOf("\\")==-1) return s; // for performace reasons: replace nothing takes up time
   var r=s.replace(/\\c/g,":");
   r=r.replace(/\\n/g,"\n");
   r=r.replace(/\\e/g,"]]");
   return r.replace(/\\b/g,"\\");
}

SocialCalc.encodeForSave = function(s) {
   if (typeof s != "string") return s;
   if (s.indexOf("\\")!=-1) // for performace reasons: replace nothing takes up time
      s=s.replace(/\\/g,"\\b");
   if (s.indexOf(":")!=-1)
      s=s.replace(/:/g,"\\c");
   if (s.indexOf("\n")!=-1)
      s=s.replace(/\n/g,"\\n");
   return s;
}

//
// Returns estring where &, <, >, " are HTML escaped
//
SocialCalc.special_chars = function(string) {

   if (/[&<>"]/.test(string)) { // only do "slow" replaces if something to replace
      string = string.replace(/&/g, "&amp;");
      string = string.replace(/</g, "&lt;");
      string = string.replace(/>/g, "&gt;");
      string = string.replace(/"/g, "&quot;");
   }
   return string;

}

SocialCalc.Lookup = function(value, list) {

   for (i=0; i<list.length; i++) {
      if (list[i] > value) {
         if (i>0) return i-1;
         else return null;
      }
   }
   return list.length-1; // if all smaller, matches last

}

//
// setStyles(element, cssText)
//
// Takes a pseudo style string (e.g., text-align must be textAlign) and sets
// the element's style value for each style name listed (leaving others unchanged).
// OK to call with null cssText.
//

SocialCalc.setStyles = function (element, cssText) {

   var parts, part, pos, name, value;

   if (!cssText) return;

   parts = cssText.split(";");
   for (part=0; part<parts.length; part++) {
      pos = parts[part].indexOf(":"); // find first colon (could be one in url)
      if (pos != -1) {
         name = parts[part].substring(0, pos);
         value = parts[part].substring(pos+1);
         if (name && value) { // if non-null name and value, set style
            element.style[name] = value;
         }
      }
//      namevalue = parts[part].split(":");
//      if (namevalue[0]) element.style[namevalue[0]] = namevalue[1];
   }

}

//
// GetViewportInfo() - returns object with viewport width and height, and scroll offsets
//
// Flanagan, JavaScript, 5th Edition, page 276
//

SocialCalc.GetViewportInfo = function () {

   var result = {};

   if (window.innerWidth) { // all but IE
      result.width = window.innerWidth;
      result.height = window.innerHeight;
      result.horizontalScroll = window.pageXOffset;
      result.verticalScroll = window.pageYOffset;
   }
   else {
      if (document.documentElement && document.documentElement.clientWidth) {
         result.width = document.documentElement.clientWidth;
         result.height = document.documentElement.clientHeight;
         result.horizontalScroll = document.documentElement.scrollLeft;
         result.verticalScroll = document.documentElement.scrollTop;
      }
      else if (document.body.clientWidth) {
         result.width = document.body.clientWidth;
         result.height = document.body.clientHeight;
         result.horizontalScroll = document.body.scrollLeft;
         result.verticalScroll = document.body.scrollTop;
      }
   }

   return result;
}

//
// GetElementPosition(element) - returns object with left and top position of the element in the document
//
// Goodman's JavaScript & DHTML Cookbook, 2nd Edition, page 415
//

SocialCalc.GetElementPosition = function (element) {

   var offsetLeft = 0;
   var offsetTop = 0;
   while (element) {
      if (SocialCalc.GetComputedStyle(element,'position')=='relative') break;
      offsetLeft+=element.offsetLeft;
      offsetTop+=element.offsetTop;
      element=element.offsetParent;
   }
   return {left:offsetLeft, top:offsetTop};

}

//
// GetElementPositionWithScroll(element) - returns object with left and top position of the element in the document
//

SocialCalc.GetElementPositionWithScroll = function (element) {

   var rect = element.getBoundingClientRect();
   return {
      left:rect.left,
      right:rect.right,
      top:rect.top,
      bottom:rect.bottom,
      width:rect.width?rect.width:rect.right-rect.left,
      height:rect.height?rect.height:rect.bottom-rect.top
   };

}

//
// GetElementFixedParent(element) - checks whether element has a parent with position:fixed
//

SocialCalc.GetElementFixedParent = function(element) {

   while (element) {
      if (element.tagName=="HTML") break;
      if (SocialCalc.GetComputedStyle(element,'position')=='fixed') return element;
      element=element.parentNode;
   }
   return false;

}

//
// GetComputedStyle(element, style) - returns computed style value
//
// http://blog.stchur.com/2006/06/21/css-computed-style/
//

SocialCalc.GetComputedStyle = function (element, style) {

   var computedStyle;
   if (typeof element.currentStyle != 'undefined') { // IE
      computedStyle = element.currentStyle;
   }
   else {
      computedStyle = document.defaultView.getComputedStyle(element, null);
   }
   return computedStyle[style];

}

//
// LookupElement(element, array) - returns array element which is an object with "element" of element
//

SocialCalc.LookupElement = function (element, array) {

   var i;
   for (i=0; i<array.length; i++) {
      if (array[i].element == element) return array[i];
   }
   return null;

}

//
// AssignID(obj, element, id) - Optionally assigns an ID with a prefix to the element
//

SocialCalc.AssignID = function (obj, element, id) {

   if (obj.idPrefix) { // Object must have a non-empty idPrefix attribute
      element.id = obj.idPrefix + id;
   }

}

//
// SocialCalc.GetCellContents(sheetobj, coord)
//
// Returns the contents (value, formula, constant, etc.) of a cell
// with appropriate prefix ("'", "=", etc.)
//

SocialCalc.GetCellContents = function(sheetobj, coord) {

   var result = "";
   var cellobj = sheetobj.cells[coord];
   if (cellobj) {
      switch (cellobj.datatype) {
         case "v":
            result = cellobj.datavalue+"";
            break;
         case "t":
            result = "'"+cellobj.datavalue;
            break;
         case "f":
            result = "="+cellobj.formula;
            break;
         case "c":
            result = cellobj.formula;
            break;
         default:
            break;
      }
   }

   return result;

}

//
// Routines translated from the SocialCalc 1.1.0 Perl code:
//
// (Makes use of the FormatNumber JavaScript code translated from the Perl.)
//

//
// displayvalue = FormatValueForDisplay(sheetobj, value, cr, linkstyle)
//
// Returns a string, in HTML, for the contents of a cell.
//
// The value is a either numeric or text, the cr is the coord of the cell
// (its cell properties are used to determine formatting), and linkstyle
// is a value passed to wiki-text expansion routines specifying the
// purpose of the rendering so, for example, links can be rendered differently
// during edit than with plain HTML.
//

SocialCalc.FormatValueForDisplay = function(sheetobj, value, cr, linkstyle) {

   var valueformat, has_parens, has_commas, valuetype, valuesubtype;
   var displayvalue;

   var sheetattribs=sheetobj.attribs;
   var scc=SocialCalc.Constants;

   var cell=sheetobj.cells[cr];

   if (!cell) { // get an empty cell if not there
      cell=new SocialCalc.Cell(cr);
   }

   displayvalue = value;

   valuetype = cell.valuetype || ""; // get type of value to determine formatting
   valuesubtype = valuetype.substring(1);
   valuetype = valuetype.charAt(0);

   if (cell.errors || valuetype=="e") {
      displayvalue = cell.errors || valuesubtype || "Error in cell";
      return displayvalue;
   }

   if (valuetype=="t") {
      valueformat = sheetobj.valueformats[cell.textvalueformat-0] || sheetobj.valueformats[sheetattribs.defaulttextvalueformat-0] || "";
      if (valueformat=="formula") {
         if (cell.datatype=="f") {
            displayvalue = SocialCalc.special_chars("="+cell.formula) || "&nbsp;";
         }
         else if (cell.datatype=="c") {
            displayvalue = SocialCalc.special_chars("'"+cell.formula) || "&nbsp;";
         }
         else {
            displayvalue = SocialCalc.special_chars("'"+displayvalue) || "&nbsp;";
         }
         return displayvalue;
      }
      displayvalue = SocialCalc.format_text_for_display(displayvalue, cell.valuetype, valueformat, sheetobj, linkstyle, cell.nontextvalueformat);
   }

   else if (valuetype=="n") {
      valueformat = cell.nontextvalueformat;
      if (valueformat==null || valueformat=="") { //
         valueformat = sheetattribs.defaultnontextvalueformat;
      }
      valueformat = sheetobj.valueformats[valueformat-0];
      if (valueformat==null || valueformat=="none") {
         valueformat = "";
      }
      if (valueformat=="formula") {
         if (cell.datatype=="f") {
            displayvalue = SocialCalc.special_chars("="+cell.formula) || "&nbsp;";
         }
         else if (cell.datatype=="c") {
            displayvalue = SocialCalc.special_chars("'"+cell.formula) || "&nbsp;";
         }
         else {
            displayvalue = SocialCalc.special_chars("'"+displayvalue) || "&nbsp;";
         }
         return displayvalue;
      }
      else if (valueformat=="forcetext") {
         if (cell.datatype=="f") {
            displayvalue = SocialCalc.special_chars("="+cell.formula) || "&nbsp;";
         }
         else if (cell.datatype=="c") {
            displayvalue = SocialCalc.special_chars(cell.formula) || "&nbsp;";
         }
         else {
            displayvalue = SocialCalc.special_chars(displayvalue) || "&nbsp;";
         }
         return displayvalue;
      }

      displayvalue = SocialCalc.format_number_for_display(displayvalue, cell.valuetype, valueformat);

   }
   else { // unknown type - probably blank
      displayvalue = "&nbsp;";
   }

   return displayvalue;

}


//
// displayvalue = format_text_for_display(rawvalue, valuetype, valueformat, sheetobj, linkstyle, nontextvalueformat)
//

SocialCalc.format_text_for_display = function(rawvalue, valuetype, valueformat, sheetobj, linkstyle, nontextvalueformat) {

   var valueformat, valuesubtype, dvsc, dvue, textval;
   var displayvalue;

   valuesubtype = valuetype.substring(1);

   displayvalue = rawvalue;

   if (valueformat=="none" || valueformat==null) valueformat="";
   if (!/^(text-|custom|hidden)/.test(valueformat)) valueformat="";
   if (valueformat=="" || valueformat=="General") { // determine format from type
      if (valuesubtype=="h") valueformat="text-html";
      if (valuesubtype=="w" || valuesubtype=="r") valueformat="text-wiki";
      if (valuesubtype=="l") valueformat="text-link";
      if (!valuesubtype) valueformat="text-plain";
   }
   if (valueformat=="text-html") { // HTML - output as it as is
      ;
   }
   else if (SocialCalc.Callbacks.expand_wiki && /^text-wiki/.test(valueformat)) { // do general wiki markup
      displayvalue = SocialCalc.Callbacks.expand_wiki(displayvalue, sheetobj, linkstyle, valueformat);
   }
   else if (valueformat=="text-wiki") { // wiki text
      displayvalue = (SocialCalc.Callbacks.expand_markup
          && SocialCalc.Callbacks.expand_markup(displayvalue, sheetobj, linkstyle)) || // do old wiki markup
          SocialCalc.special_chars("wiki-text:"+displayvalue);
   }
   else if (valueformat=="text-url") { // text is a URL for a link
      dvsc = SocialCalc.special_chars(displayvalue);
      dvue = encodeURI(displayvalue);
      displayvalue = '<a href="'+dvue+'">'+dvsc+'</a>';
   }
   else if (valueformat=="text-link") { // more extensive link capabilities for regular web links
      displayvalue = SocialCalc.expand_text_link(displayvalue, sheetobj, linkstyle, valueformat);
   }
   else if (valueformat=="text-image") { // text is a URL for an image
      dvue = encodeURI(displayvalue);
      displayvalue = '<img src="'+dvue+'">';
   }
   else if (valueformat.substring(0,12)=="text-custom:") { // construct a custom text format: @r = text raw, @s = special chars, @u = url encoded
      dvsc = SocialCalc.special_chars(displayvalue); // do special chars
      dvsc = dvsc.replace(/  /g, "&nbsp; "); // keep multiple spaces
      dvsc = dvsc.replace(/\n/g, "<br>");  // keep line breaks
      dvue = encodeURI(displayvalue);
      textval={};
      textval.r = displayvalue;
      textval.s = dvsc;
      textval.u = dvue;
      displayvalue = valueformat.substring(12); // remove "text-custom:"
      displayvalue = displayvalue.replace(/@(r|s|u)/g, function(a,c){return textval[c];}); // replace placeholders
   }
   else if (valueformat.substring(0,6)=="custom") { // custom
      displayvalue = SocialCalc.special_chars(displayvalue); // do special chars
      displayvalue = displayvalue.replace(/  /g, "&nbsp; "); // keep multiple spaces
      displayvalue = displayvalue.replace(/\n/g, "<br>"); // keep line breaks
      displayvalue += " (custom format)";
   }
   else if (valueformat=="hidden") {
      displayvalue = "&nbsp;";
   }
   else if (nontextvalueformat != null && nontextvalueformat != "" && sheetobj.valueformats[nontextvalueformat-0] != "none" && sheetobj.valueformats[nontextvalueformat-0] != "" ) {
      valueformat = sheetobj.valueformats[nontextvalueformat];
      displayvalue = SocialCalc.format_number_for_display(rawvalue, valuetype, valueformat);
   }
   else { // plain text
      displayvalue = SocialCalc.special_chars(displayvalue); // do special chars
      displayvalue = displayvalue.replace(/  /g, "&nbsp; "); // keep multiple spaces
      displayvalue = displayvalue.replace(/\n/g, "<br>"); // keep line breaks
   }

   return displayvalue;

}


//
// displayvalue = format_number_for_display(rawvalue, valuetype, valueformat)
//

SocialCalc.format_number_for_display = function(rawvalue, valuetype, valueformat) {

   var value, valuesubtype;
   var scc = SocialCalc.Constants;

   value = rawvalue-0;

   valuesubtype = valuetype.substring(1);

   if (valueformat=="Auto" || valueformat=="") { // cases with default format
      if (valuesubtype=="%") { // will display a % character
         valueformat = scc.defaultFormatp;
      }
      else if (valuesubtype=='$') {
         valueformat = scc.defaultFormatc;
      }
      else if (valuesubtype=='dt') {
         valueformat = scc.defaultFormatdt;
      }
      else if (valuesubtype=='d') {
         valueformat = scc.defaultFormatd;
      }
      else if (valuesubtype=='t') {
         valueformat = scc.defaultFormatt;
      }
      else if (valuesubtype=='l') {
         valueformat = 'logical';
      }
      else {
         valueformat = "General";
      }
   }

   if (valueformat=="logical") { // do logical format
      return value ? scc.defaultDisplayTRUE : scc.defaultDisplayFALSE;
   }

   if (valueformat=="hidden") { // do hidden format
      return "&nbsp;";
   }

   // Use format

   return SocialCalc.FormatNumber.formatNumberWithFormat(rawvalue, valueformat, "");

}

//
// valueinfo = DetermineValueType(rawvalue)
//
// Takes a value and looks for special formatting like $, %, numbers, etc.
// Returns the value as a number or string and the type as {value: value, type: type}.
// Tries to follow the spec for spreadsheet function VALUE(v).
//

SocialCalc.DetermineValueType = function(rawvalue) {

   var value = rawvalue + "";
   var type = "t";
   var tvalue, matches, year, hour, minute, second, denom, num, intgr, constr;

   tvalue = value.replace(/^\s+/, ""); // remove leading and trailing blanks
   tvalue = tvalue.replace(/\s+$/, "");

   if (value.length==0) {
      type = "";
   }
   else if (value.match(/^\s+$/)) { // just blanks
      ; // leave type "t"
   }
   else if (tvalue.match(/^[-+]?\d*(?:\.)?\d*(?:[eE][-+]?\d+)?$/)) { // general number, including E
      value = tvalue - 0; // try converting to number
      if (isNaN(value)) { // leave alone - catches things like plain "-"
         value = rawvalue + "";
      }
      else {
         type = "n";
      }
   }
   else if (tvalue.match(/^[-+]?\d*(?:\.)?\d*\s*%$/)) { // percent form: 15.1%
      value = (tvalue.slice(0, -1) - 0) / 100; // convert and scale
      type = "n%";
   }
   else if (tvalue.match(/^[-+]?\$\s*\d*(?:\.)?\d*\s*$/) && tvalue.match(/\d/)) { // $ format: $1.49
      value = tvalue.replace(/\$/, "") - 0;
      type = "n$";
   }
   else if (tvalue.match(/^[-+]?(\d*,\d*)+(?:\.)?\d*$/)) { // number format ignoring commas: 1,234.49
      value = tvalue.replace(/,/g, "") - 0;
      type = "n";
   }
   else if (tvalue.match(/^[-+]?(\d*,\d*)+(?:\.)?\d*\s*%$/)) { // % with commas: 1,234.49%
      value = (tvalue.replace(/[%,]/g, "") - 0) / 100;
      type = "n%";
   }
   else if (tvalue.match(/^[-+]?\$\s*(\d*,\d*)+(?:\.)?\d*$/) && tvalue.match(/\d/)) { // $ and commas: $1,234.49
      value = tvalue.replace(/[\$,]/g, "") - 0;
      type = "n$";
   }
   else if (matches=value.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{1,4})\s*$/)) { // MM/DD/YYYY, MM/DD/YYYY
      year = matches[3] - 0;
      year = year < 1000 ? year + 2000 : year;
      value = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(year, matches[1]-0, matches[2]-0)-2415019;
      type = "nd";
   }
   else if (matches=value.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s*$/)) { // YYYY-MM-DD, YYYY/MM/DD
      year = matches[1]-0;
      year = year < 1000 ? year + 2000 : year;
      value = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(year, matches[2]-0, matches[3]-0)-2415019;
      type = "nd";
   }
   else if (matches=value.match(/^(\d{1,2}):(\d{1,2})\s*$/)) { // HH:MM
      hour = matches[1]-0;
      minute = matches[2]-0;
      if (hour < 24 && minute < 60) {
         value = hour/24 + minute/(24*60);
         type = "nt";
      }
   }
   else if (matches=value.match(/^(\d{1,2}):(\d{1,2}):(\d{1,2})\s*$/)) { // HH:MM:SS
      hour = matches[1]-0;
      minute = matches[2]-0;
      second = matches[3]-0;
      if (hour < 24 && minute < 60 && second < 60) {
         value = hour/24 + minute/(24*60) + second/(24*60*60);
         type = "nt";
      }
   }
   else if (matches=value.match(/^\s*([-+]?\d+) (\d+)\/(\d+)\s*$/)) { // 1 1/2
      intgr = matches[1]-0;
      num = matches[2]-0;
      denom = matches[3]-0;
      if (denom && denom > 0) {
         value = intgr + (intgr < 0 ? -num/denom : num/denom);
         type = "n";
      }
   }
   else if (constr=SocialCalc.InputConstants[value.toUpperCase()]) { // special constants, like "false" and #N/A
      num = constr.indexOf(",");
      value = constr.substring(0,num)-0;
      type = constr.substring(num+1);
   }
   else if (tvalue.length > 7 && tvalue.substring(0,7).toLowerCase()=="http://") { // URL
      value = tvalue;
      type = "tl";
   }
   else if (tvalue.match(/<([A-Z][A-Z0-9]*)\b[^>]*>[\s\S]*?<\/\1>/i)) { // HTML
      value = tvalue;
      type = "th";
   }

   return {value: value, type: type};

}

SocialCalc.InputConstants = { // strings that turn into constants for SocialCalc.DetermineValueType
   "TRUE": "1,nl", "FALSE": "0,nl", "#N/A": "0,e#N/A", "#NULL!": "0,e#NULL!", "#NUM!": "0,e#NUM!",
   "#DIV/0!": "0,e#DIV/0!", "#VALUE!": "0,e#VALUE!", "#REF!": "0,e#REF!", "#NAME?": "0,e#NAME?"};

//
// result = default_expand_markup(displayvalue, sheetobj, linkstyle)
//
// Processes wiki-text -- this is a placeholder.
// Reference to here in SocialCalc.expand_markup should be replaced by application-specific routine.
//

SocialCalc.default_expand_markup = function(displayvalue, sheetobj, linkstyle) {

   var result = displayvalue;

   result = SocialCalc.special_chars(result); // do special chars
   result = result.replace(/  /g, "&nbsp; "); // keep multiple spaces
   result = result.replace(/\n/g, "<br>"); // keep line breaks

   return result; // do very little by default

   result = result.replace(/('*)'''(.*?)'''/g, "$1<b>$2<\/b>"); // Wiki-style bold/italics
   result = result.replace(/''(.*?)''/g, "<i>$1<\/i>");

   return result;

}


//
// result = SocialCalc.expand_text_link(displayvalue, sheetobj, linkstyle, valueformat)
//
// Parses link text (URL, descriptions, pagenames, workspace names) and returns HTML
//

SocialCalc.expand_text_link = function(displayvalue, sheetobj, linkstyle, valueformat) {

   var desc, tb, str;

   var scc = SocialCalc.Constants;

   var url = "";
   var parts = SocialCalc.ParseCellLinkText(displayvalue+"");

   if (parts.desc) {
      desc = SocialCalc.special_chars(parts.desc);
   }
   else {
      desc = parts.pagename ? scc.defaultPageLinkFormatString : scc.defaultLinkFormatString;
   }

   if (displayvalue.length > 7 && displayvalue.substring(0,7).toLowerCase()=="http://"
       && displayvalue.charAt(displayvalue.length-1)!=">") {
      desc = desc.substring(7); // remove http:// unless explicit
   }

   tb = (parts.newwin || !linkstyle) ? ' target="_blank"' : "";

   if (parts.pagename) {
      if (SocialCalc.Callbacks.MakePageLink) {
         url = SocialCalc.Callbacks.MakePageLink(parts.pagename, parts.workspacename, linkstyle, valueformat);
      }
//      else if (parts.workspace) {
//         url = "/" + encodeURI(parts.workspace) + "/" + encodeURI(parts.pagename);
//         }
//      else {
//         url = parts.pagename;
//         }
   }
   else {
      url = encodeURI(parts.url);
   }
   str = '<a href="' + url + '"' + tb + '>' + desc + '</a>';

   return str;

}


//
// result = SocialCalc.ParseCellLinkText(str)
//
// Given: url = http://www.someurl.com/more, desc = Some descriptive text
//
// Takes the following:
//
//    url
//    <url>
//    desc<url>
//    "desc"<url>
//    <<>> instead of <> => target="_blank" (new window)
//
//    [page name]
//    "desc"[page name]
//    desc[page name]
//    {workspace name [page name]}
//    "desc"{workspace name [page name]}
//    [[]] instead of [] => target="_blank" (new window)
//
//
// Returns: {url: url, desc: desc, newwin: t/f, pagename: pagename, workspace: workspace}
//

SocialCalc.ParseCellLinkText = function(str) {

   var result = {url: "", desc: "", newwin: false, pagename: "", workspace: ""};

   var pageform = false;
   var urlend = str.length - 1;
   var descstart = 0;
   var lastlt = str.lastIndexOf("<");
   var lastbrkt = str.lastIndexOf("[");
   var lastbrace = str.lastIndexOf("{");
   var descend = -1;

   if ((str.charAt(urlend) != ">" || lastlt == -1)
       && (str.charAt(urlend) != "]" || lastbrkt == -1)
       && (str.charAt(urlend) != "}" || str.charAt(urlend-1) != "]" ||
       lastbrace == -1 || lastbrkt == -1 || lastbrkt < lastbrace)) { // plain url
      urlend++;
      descend = urlend;
   }
   else { // some markup
      if (str.charAt(urlend)==">") { // url form
         descend = lastlt - 1;
         if (lastlt > 0 && str.charAt(descend) == "<" && str.charAt(urlend-1) == ">") {
            descend--;
            urlend--;
            result.newwin = true;
         }
      }

      else if (str.charAt(urlend)=="]") { // plain page form
         descend = lastbrkt - 1;
         pageform = true;
         if (lastbrkt > 0 && str.charAt(descend) == "[" && str.charAt(urlend-1) == "]") {
            descend--;
            urlend--;
            result.newwin = true;
         }
      }

      else if (str.charAt(urlend)=="}") { // page and workspace form
         descend = lastbrace - 1;
         pageform = true;
         wsend = lastbrkt;
         urlend--;
         if (lastbrkt > 0 && str.charAt(lastbrkt-1) == "[" && str.charAt(urlend-1) == "]") {
            wsend = lastbrkt-1;
            urlend--;
            result.newwin = true;
         }
         if (str.charAt(wsend-1)==" ") { // trim trailing space in workspace name
            wsend--;
         }
         result.workspace = str.substring(lastbrace+1, wsend) || "";
      }

      if (str.charAt(descend)==" ") { // trim trailing space on desc
         descend--;
      }

      if (str.charAt(descstart) == '"' && str.charAt(descend) == '"') {
         descstart++;
         descend--;
      }
   }

   if (pageform) {
      result.pagename = str.substring(lastbrkt+1, urlend) || "";
   }
   else {
      result.url = str.substring(lastlt+1, urlend) || "";
   }

   if (descend >= descstart) {
      result.desc = str.substring(descstart, descend+1);
   }

   return result;

}


//
// result = SocialCalc.ConvertSaveToOtherFormat(savestr, outputformat, dorecalc)
//
// Returns a string in the specificed format: "scsave", "html", "csv", "tab" (tab delimited)
// If dorecalc is true, performs a recalc after loading (NO: obsolete!).
//

SocialCalc.ConvertSaveToOtherFormat = function(savestr, outputformat, dorecalc) {

   var sheet, context, clipextents, div, ele, row, col, cr, cell, str;

   var result = "";

   if (outputformat == "scsave") {
      return savestr;
   }

   if (savestr == "") {
      return "";
   }

   sheet = new SocialCalc.Sheet();
   sheet.ParseSheetSave(savestr);

   if (dorecalc) {
      // no longer supported as of 9/10/08
      // Recalc is now async, so can't do it this way
      throw("SocialCalc.ConvertSaveToOtherFormat: Not doing recalc.");
   }

   if (sheet.copiedfrom) {
      clipextents = SocialCalc.ParseRange(sheet.copiedfrom);
   }
   else {
      clipextents = {cr1: {row: 1, col: 1}, cr2: {row: sheet.attribs.lastrow, col: sheet.attribs.lastcol}};
   }

   if (outputformat == "html") {
      context=new SocialCalc.RenderContext(sheet);
      if (sheet.copiedfrom) {
         context.rowpanes[0] = {first: clipextents.cr1.row, last: clipextents.cr2.row};
         context.colpanes[0] = {first: clipextents.cr1.col, last: clipextents.cr2.col};
      }
      div = document.createElement("div");
      ele = context.RenderSheet(null, context.defaultHTMLlinkstyle);
      div.appendChild(ele);
      context = null;
      sheet = null;
      result = div.innerHTML;
      ele = null;
      div = null;
      return result;
   }

   for (row = clipextents.cr1.row; row <= clipextents.cr2.row; row++) {
      for (col = clipextents.cr1.col; col <= clipextents.cr2.col; col++) {
         cr = SocialCalc.crToCoord(col, row);
         cell = sheet.GetAssuredCell(cr);

         if (cell.errors) {
            str = cell.errors;
         }
         else {
            str = cell.datavalue+""; // get value as text
         }

         if (outputformat == "csv") {
            if (str.indexOf('"')!=-1) {
               str = str.replace(/"/g, '""'); // double quotes
            }
            if (/[, \n"]/.test(str)) {
               str = '"' + str + '"'; // add quotes
            }
            if (col>clipextents.cr1.col) {
               str = "," + str; // add commas
            }
         }
         else if (outputformat == "tab") {
            if (str.indexOf('\n')!=-1) { // if multiple lines
               if (str.indexOf('"')!=-1) {
                  str = str.replace(/"/g, '""'); // double quotes
               }
               str = '"' + str + '"'; // add quotes
            }
            if (col>clipextents.cr1.col) {
               str = "\t" + str; // add tabs
            }
         }
         result += str;
      }
      result += "\n";
   }

   return result;

}


//
// result = SocialCalc.ConvertOtherFormatToSave(inputstr, inputformat)
//
// Returns a string converted from the specified format: "scsave", "csv", "tab" (tab delimited)
//

SocialCalc.ConvertOtherFormatToSave = function(inputstr, inputformat) {

   var sheet, context, lines, i, line, value, inquote, j, ch, values, row, col, cr, maxc;

   var result = "";

   var AddCell = function() {
      col++;
      if (col>maxc) maxc = col;
      cr = SocialCalc.crToCoord(col, row);
      SocialCalc.SetConvertedCell(sheet, cr, value);
      value = "";
   }

   if (inputformat == "scsave") {
      return inputstr;
   }

   sheet = new SocialCalc.Sheet();

   lines = inputstr.split(/\r\n|\n/);

   maxc = 0;
   if (inputformat == "csv") {
      row = 0;
      inquote = false;
      for (i=0; i<lines.length; i++) {
         if (i==lines.length-1 && lines[i]=="") { // extra null line - ignore
            break;
         }
         if (inquote) { // if inquote, just continue from where left off
            value += "\n";
         }
         else { // otherwise next row
            value = "";
            row++;
            col = 0;
         }
         line = lines[i];
         for (j=0; j<line.length; j++) {
            ch = line.charAt(j);
            if (ch == '"') {
               if (inquote) {
                  if (j<line.length-1 && line.charAt(j+1) == '"') { // double quotes
                     j++; // skip the second one
                     value += '"'; // add one quote
                  }
                  else {
                     inquote = false;
                     if (j==line.length-1) { // at end of line
                        AddCell();
                     }
                  }
               }
               else {
                  inquote = true;
               }
               continue;
            }
            if (ch == "," && !inquote) {
               AddCell();
            }
            else {
               value += ch;
            }
            if (j==line.length-1 && !inquote) {
               AddCell();
            }
         }
      }
      if (maxc>0) {
         sheet.attribs.lastrow = row;
         sheet.attribs.lastcol = maxc;
         result = sheet.CreateSheetSave("A1:"+SocialCalc.crToCoord(maxc, row));
      }
   }

   if (inputformat == "tab") {
      row = 0;
      inquote = false;
      for (i=0; i<lines.length; i++) {
         if (i==lines.length-1 && lines[i]=="") { // extra null line - ignore
            break;
         }
         if (inquote) { // if inquote, just continue from where left off
            value += "\n";
         }
         else { // otherwise next row
            value = "";
            row++;
            col = 0;
         }
         line = lines[i];
         for (j=0; j<line.length; j++) {
            ch = line.charAt(j);
            if (ch == '"') {
               if (inquote) {
                  if (j<line.length-1) {
                     if (line.charAt(j+1) == '"') { // double quotes
                        j++; // skip the second one
                        value += '"'; // add one quote
                     }
                     else if (line.charAt(j+1) == '\t') { // end of quoted item
                        j++;
                        inquote = false;
                        AddCell();
                     }
                  }
                  else { // at end of line
                     inquote = false;
                     AddCell();
                  }
                  continue;
               }
               if (value=="") { // quote at start of item
                  inquote = true;
                  continue;
               }
            }
            if (ch == "\t" && !inquote) {
               AddCell();
            }
            else {
               value += ch;
            }
            if (j==line.length-1 && !inquote) {
               AddCell();
            }
         }
      }
      if (maxc>0) {
         sheet.attribs.lastrow = row;
         sheet.attribs.lastcol = maxc;
         result = sheet.CreateSheetSave("A1:"+SocialCalc.crToCoord(maxc, row));
      }
   }

   return result;

}

//
// SocialCalc.SetConvertedCell(sheet, cr, rawvalue)
//
// Sets the cell cr with a value and type determined from rawvalue
//

SocialCalc.SetConvertedCell = function(sheet, cr, rawvalue) {

   var cell, value;

   cell = sheet.GetAssuredCell(cr);

   value = SocialCalc.DetermineValueType(rawvalue);

   if (value.type == 'n' && value.value == rawvalue) { // check that we don't need "constant" to remember original value
      cell.datatype = "v";
      cell.valuetype = "n";
      cell.datavalue = value.value;
   }
   else if (value.type.charAt(0) == 't') { // text of some sort but left unchanged
      cell.datatype = "t";
      cell.valuetype = value.type;
      cell.datavalue = value.value;
   }
   else { // special number types
      cell.datatype = "c";
      cell.valuetype = value.type;
      cell.datavalue = value.value;
      cell.formula = rawvalue;
   }

}

//
/*
// SocialCalc Spreadsheet Formula Library
//
// Part of the SocialCalc package
//
// (c) Copyright 2008 Socialtext, Inc.
// All Rights Reserved.
//
// The contents of this file are subject to the Artistic License 2.0; you may not
// use this file except in compliance with the License. You may obtain a copy of 
// the License at http://socialcalc.org/licenses/al-20/.
//
// Some of the other files in the SocialCalc package are licensed under
// different licenses. Please note the licenses of the modules you use.
//
// Code History:
//
// Initially coded by Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
// Based in part on the SocialCalc 1.1.0 code written in Perl.
// The SocialCalc 1.1.0 code was:
//    Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc.
//    All Rights Reserved.
//    Portions (c) Copyright 2007 Socialtext, Inc.
//    All Rights Reserved.
// The Perl SocialCalc started as modifications to the wikiCalc(R) program, version 1.0.
// wikiCalc 1.0 was written by Software Garden, Inc.
// Unless otherwise specified, referring to "SocialCalc" in comments refers to this
// JavaScript version of the code, not the SocialCalc Perl code.
//
*/

   var SocialCalc;
   if (!SocialCalc) SocialCalc = {}; // May be used with other SocialCalc libraries or standalone
                                     // In any case, requires SocialCalc.Constants.

SocialCalc.Formula = {};

//
// Formula constants for parsing:
//

   SocialCalc.Formula.ParseState = {num: 1, alpha: 2, coord: 3, string: 4, stringquote: 5, numexp1: 6, numexp2: 7, alphanumeric: 8, specialvalue:9};

   SocialCalc.Formula.TokenType = {num: 1, coord: 2, op: 3, name: 4, error: 5, string: 6, space: 7};

   SocialCalc.Formula.CharClass = {num: 1, numstart: 2, op: 3, eof: 4, alpha: 5, incoord: 6, error: 7, quote: 8, space: 9, specialstart: 10};
 
   SocialCalc.Formula.CharClassTable = {
      " ": 9, "!": 3, '"': 8, "'": 8, "#": 10, "$":6, "%":3, "&":3, "(": 3, ")": 3, "*": 3, "+": 3, ",": 3, "-": 3, ".": 2, "/": 3,
       "0": 1, "1": 1, "2": 1, "3": 1, "4": 1, "5": 1, "6": 1, "7": 1, "8": 1, "9": 1,
       ":": 3, "<": 3, "=": 3, ">": 3,
       "A": 5, "B": 5, "C": 5, "D": 5, "E": 5, "F": 5, "G": 5, "H": 5, "I": 5, "J": 5, "K": 5, "L": 5, "M": 5, "N": 5,
       "O": 5, "P": 5, "Q": 5, "R": 5, "S": 5, "T": 5, "U": 5, "V": 5, "W": 5, "X": 5, "Y": 5, "Z": 5,
       "^": 3, "_": 5,
       "a": 5, "b": 5, "c": 5, "d": 5, "e": 5, "f": 5, "g": 5, "h": 5, "i": 5, "j": 5, "k": 5, "l": 5, "m": 5, "n": 5,
       "o": 5, "p": 5, "q": 5, "r": 5, "s": 5, "t": 5, "u": 5, "v": 5, "w": 5, "x": 5, "y": 5, "z": 5
       };

   SocialCalc.Formula.UpperCaseTable = {
       "a": "A", "b": "B", "c": "C", "d": "D", "e": "E", "f": "F", "g": "G", "h": "H", "i": "I", "j": "J", "k": "K", "l": "L", "m": "M",
       "n": "N", "o": "O", "p": "P", "q": "Q", "r": "R", "s": "S", "t": "T", "u": "U", "v": "V", "w": "W", "x": "X", "y": "Y", "z": "Z",
       "A": "A", "B": "B", "C": "C", "D": "D", "E": "E", "F": "F", "G": "G", "H": "H", "I": "I", "J": "J", "K": "K", "L": "L", "M": "M",
       "N": "N", "O": "O", "P": "P", "Q": "Q", "R": "R", "S": "S", "T": "T", "U": "U", "V": "V", "W": "W", "X": "X", "Y": "Y", "Z": "Z"
       }

   SocialCalc.Formula.SpecialConstants = { // names that turn into constants for name lookup
      "#NULL!": "0,e#NULL!", "#NUM!": "0,e#NUM!", "#DIV/0!": "0,e#DIV/0!", "#VALUE!": "0,e#VALUE!",
      "#REF!": "0,e#REF!", "#NAME?": "0,e#NAME?"};


   // Operator Precedence table
   //
   // 1- !, 2- : ,, 3- M P, 4- %, 5- ^, 6- * /, 7- + -, 8- &, 9- < > = G(>=) L(<=) N(<>),
   // Negative value means Right Associative

   SocialCalc.Formula.TokenPrecedence = {
      "!": 1,
      ":": 2, ",": 2,
      "M": -3, "P": -3,
      "%": 4,
      "^": 5,
      "*": 6, "/": 6,
      "+": 7, "-": 7,
      "&": 8,
      "<": 9, ">": 9, "G": 9, "L": 9, "N": 9
      };

   // Convert one-char token text to input text:

   SocialCalc.Formula.TokenOpExpansion = {'G': '>=', 'L': '<=', 'M': '-', 'N': '<>', 'P': '+'};

   //
   // Information about the resulting value types when doing operations on values (used by LookupResultType)
   //
   // Each object entry is an object with specific types with result type info as follows:
   //
   //    'type1a': '|type2a:resulta|type2b:resultb|...
   //    Type of t* or n* matches any of those types not listed
   //    Results may be a type or the numbers 1 or 2 specifying to return type1 or type2
   

   SocialCalc.Formula.TypeLookupTable = {
       unaryminus: { 'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
       unaryplus: { 'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
       unarypercent: { 'n*': '|n:n%|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
       plus: {
                'n%': '|n%:n%|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
                'nd': '|n%:n|nd:nd|nt:ndt|ndt:ndt|n$:n|n:nd|n*:n|b:n|e*:2|t*:e#VALUE!|',
                'nt': '|n%:n|nd:ndt|nt:nt|ndt:ndt|n$:n|n:nt|n*:n|b:n|e*:2|t*:e#VALUE!|',
                'ndt': '|n%:n|nd:ndt|nt:ndt|ndt:ndt|n$:n|n:ndt|n*:n|b:n|e*:2|t*:e#VALUE!|',
                'n$': '|n%:n|nd:n|nt:n|ndt:n|n$:n$|n:n$|n*:n|b:n|e*:2|t*:e#VALUE!|',
                'nl': '|n%:n|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
                'n': '|n%:n|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
                'b': '|n%:n%|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
                't*': '|n*:e#VALUE!|t*:e#VALUE!|b:e#VALUE!|e*:2|',
                'e*': '|e*:1|n*:1|t*:1|b:1|'
               },
       concat: {
                't': '|t:t|th:th|tw:tw|tl:t|tr:tr|t*:2|e*:2|',
                'th': '|t:th|th:th|tw:t|tl:th|tr:t|t*:t|e*:2|',
                'tw': '|t:tw|th:t|tw:tw|tl:tw|tr:tw|t*:t|e*:2|',
                'tl': '|t:tl|th:th|tw:tw|tl:tl|tr:tr|t*:t|e*:2|',
                't*': '|t*:t|e*:2|',
                'e*': '|e*:1|n*:1|t*:1|'
               },
       oneargnumeric: { 'n*': '|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
       twoargnumeric: { 'n*': '|n*:n|t*:e#VALUE!|e*:2|', 'e*': '|e*:1|n*:1|t*:1|', 't*': '|t*:e#VALUE!|n*:e#VALUE!|e*:2|'},
       propagateerror: { 'n*': '|n*:2|e*:2|', 'e*': '|e*:2|', 't*': '|t*:2|e*:2|', 'b': '|b:2|e*:2|'}
      };

/* *******************

 parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(line)

 Parses a text string as if it was a spreadsheet formula

 This uses a simple state machine run on each character in turn.
 States remember whether a number is being gathered, etc.
 The result is parseinfo which is an array with one entry for each token:
   parseinfo[i] = {
     text: "the characters making up the parsed token",
     type: the type of the token (a number),
     opcode: a single character version of an operator suitable for use in the
                  precedence table and distinguishing between unary and binary + and -.

************************* */

SocialCalc.Formula.ParseFormulaIntoTokens = function(line) {

   var i, ch, cclass, haddecimal, last_token, last_token_type, last_token_text, t;

   var scf = SocialCalc.Formula;
   var scc = SocialCalc.Constants;
   var parsestate = scf.ParseState;
   var tokentype = scf.TokenType;
   var charclass = scf.CharClass;
   var charclasstable = scf.CharClassTable;
   var uppercasetable = scf.UpperCaseTable; // much faster than toUpperCase function
   var pushtoken = scf.ParsePushToken;
   var coordregex = /^\$?[A-Z]{1,2}\$?[1-9]\d*$/i;

   var parseinfo = [];
   var str = "";
   var state = 0;
   var haddecimal = false;

  for (i=0; i<=line.length; i++) {
      if (i<line.length) {
         ch = line.charAt(i);
         cclass = charclasstable[ch];
         }
      else {
         ch = "";
         cclass = charclass.eof;
         }

      if (state == parsestate.num) {
         if (cclass == charclass.num) {
            str += ch;
            }
         else if (cclass == charclass.numstart && !haddecimal) {
            haddecimal = true;
            str += ch;
            }
         else if (ch == "E" || ch == "e") {
            str += ch;
            haddecimal = false;
            state = parsestate.numexp1;
            }
         else { // end of number - save it
            pushtoken(parseinfo, str, tokentype.num, 0);
            haddecimal = false;
            state = 0;
            }
         }

      if (state == parsestate.numexp1) {
         if (cclass == parsestate.num) {
            state = parsestate.numexp2;
            }
         else if ((ch == '+' || ch == '-') && (uppercasetable[str.charAt(str.length-1)] == 'E')) {
            str += ch;
            }
         else if (ch == 'E' || ch == 'e') {
            ;
            }
         else {
            pushtoken(parseinfo, scc.s_parseerrexponent, tokentype.error, 0);
            state = 0;
            }
         }

      if (state == parsestate.numexp2) {
         if (cclass == charclass.num) {
            str += ch;
            }
         else { // end of number - save it
            pushtoken(parseinfo, str, tokentype.num, 0);
            state = 0;
            }
         }

      if (state == parsestate.alpha) {
         if (cclass == charclass.num) {
            state = parsestate.coord;
            }
         else if (cclass == charclass.alpha || ch == ".") { // alpha may be letters, numbers, "_", or "."
            str += ch;
            }
         else if (cclass == charclass.incoord) {
            state = parsestate.coord;
            }
         else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
            pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
            state = 0;
            }
         else {
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            state = 0;
            }
         }

      if (state == parsestate.coord) {
         if (cclass == charclass.num) {
            str += ch;
            }
         else if (cclass == charclass.incoord) {
            str += ch;
            }
         else if (cclass == charclass.alpha) {
            state = parsestate.alphanumeric;
            }
         else if (cclass == charclass.op || cclass == charclass.numstart ||
                  cclass == charclass.eof || cclass == charclass.space) {
            if (coordregex.test(str)) {
               t = tokentype.coord;
               }
            else {
               t = tokentype.name;
               }
            pushtoken(parseinfo, str.toUpperCase(), t, 0);
            state = 0;
            }
         else {
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            state = 0;
           }
         }


      if (state == parsestate.alphanumeric) {
         if (cclass == charclass.num || cclass == charclass.alpha) {
            str += ch;
            }
         else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
            pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
            state = 0;
            }
         else {
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            state = 0;
            }
         }

      if (state == parsestate.string) {
         if (cclass == charclass.quote) {
            state = parsestate.stringquote; // got quote in string: is it doubled (quote in string) or by itself (end of string)?
            }
         else if (cclass == charclass.eof) {
            pushtoken(parseinfo, scc.s_parseerrstring, tokentype.error, 0);
            state = 0;
            }
         else {
            str += ch;
            }
         }
      else if (state == parsestate.stringquote) { // note else if here
         if (cclass == charclass.quote) {
            str += ch;
            state = parsestate.string; // double quote: add one then continue getting string
            }
         else { // something else -- end of string
            pushtoken(parseinfo, str, tokentype.string, 0);
            state = 0; // drop through to process
            }
         }

      else if (state == parsestate.specialvalue) { // special values like #REF!
         if (str.charAt(str.length-1) == "!") { // done - save value as a name
            pushtoken(parseinfo, str, tokentype.name, 0);
            state = 0; // drop through to process
            }
         else if (cclass == charclass.eof) {
            pushtoken(parseinfo, scc.s_parseerrspecialvalue, tokentype.error, 0);
            state = 0;
            }
         else {
            str += ch;
            }
         }

      if (state == 0) {
         if (cclass == charclass.num) {
            str = ch;
            state = parsestate.num;
            }
         else if (cclass == charclass.numstart) {
            str = ch;
            haddecimal = true;
            state = parsestate.num;
            }
         else if (cclass == charclass.alpha || cclass == charclass.incoord) {
            str = ch;
            state = parsestate.alpha;
            }
         else if (cclass == charclass.specialstart) {
            str = ch;
            state = parsestate.specialvalue;
            }
         else if (cclass == charclass.op) {
            str = ch;
            if (parseinfo.length>0) {
               last_token = parseinfo[parseinfo.length-1];
               last_token_type = last_token.type;
               last_token_text = last_token.text;
               if (last_token_type == charclass.op) {
                  if (last_token_text == '<' || last_token_text == ">") {
                     str = last_token_text + str;
                     parseinfo.pop();
                     if (parseinfo.length>0) {
                        last_token = parseinfo[parseinfo.length-1];
                        last_token_type = last_token.type;
                        last_token_text = last_token.text;
                        }
                     else {
                        last_token_type = charclass.eof;
                        last_token_text = "EOF";
                        }
                     }
                  }
               }
            else {
               last_token_type = charclass.eof;
               last_token_text = "EOF";
               }
            t = tokentype.op;
            if ((parseinfo.length == 0)
                || (last_token_type == charclass.op && last_token_text != ')' && last_token_text != '%')) { // Unary operator
               if (str == '-') { // M is unary minus
                  str = "M";
                  ch = "M";
                  }
               else if (str == '+') { // P is unary plus
                  str = "P";
                  ch = "P";
                  }
               else if (str == ')' && last_token_text == '(') { // null arg list OK
                  ;
                  }
               else if (str != '(') { // binary-op open-paren OK, others no
                  t = tokentype.error;
                  str = scc.s_parseerrtwoops;
                  }
               }
            else if (str.length > 1) {
               if (str == '>=') { // G is >=
                  str = "G";
                  ch = "G";
                  }
               else if (str == '<=') { // L is <=
                  str = "L";
                  ch = "L";
                  }
               else if (str == '<>') { // N is <>
                  str = "N";
                  ch = "N";
                  }
               else {
                  t = tokentype.error;
                  str = scc.s_parseerrtwoops;
                  }
               }
            pushtoken(parseinfo, str, t, ch);
            state = 0;
            }
         else if (cclass == charclass.quote) { // starting a string
            str = "";
            state = parsestate.string;
            }
         else if (cclass == charclass.space) { // store so can reconstruct spacing
            //pushtoken(parseinfo, " ", tokentype.space, 0);
            }
         else if (cclass == charclass.eof) { // ignore -- needed to have extra loop to close out other things
            }
         else { // unknown class - such as unknown char
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            }
         }
      }

   return parseinfo;

   }

SocialCalc.Formula.ParsePushToken = function(parseinfo, ttext, ttype, topcode) {

   parseinfo.push({text: ttext, type: ttype, opcode: topcode});

   }

/* *******************

 result = SocialCalc.Formula.evaluate_parsed_formula(parseinfo, sheet, allowrangereturn)

 Does the calculation expressed in a parsed formula, returning a value, its type, and error info
 returns: {value: value, type: valuetype, error: errortext}.

 If allowrangereturn is present and true, can return a range (e.g., "A1:A10" - translated from "A1|A10|")

************************* */

SocialCalc.Formula.evaluate_parsed_formula = function(parseinfo, sheet, allowrangereturn) {

   var result;

   var scf = SocialCalc.Formula;
   var tokentype = scf.TokenType;

   var revpolish;
   var parsestack = [];

   var errortext = "";

   revpolish = scf.ConvertInfixToPolish(parseinfo); // result is either an array or a string with error text

   result = scf.EvaluatePolish(parseinfo, revpolish, sheet, allowrangereturn);

   return result;

}

//
// revpolish = SocialCalc.Formula.ConvertInfixToPolish(parseinfo)
//
// Convert infix to reverse polish notation
//
// Returns revpolish array with a sequence of references to tokens by number if successful.
// Errors return a string with the error.
//
// Based upon the algorithm shown in Wikipedia "Reverse Polish notation" article
// and then enhanced for additional spreadsheet things
//

SocialCalc.Formula.ConvertInfixToPolish = function(parseinfo) {

   var scf = SocialCalc.Formula;
   var scc = SocialCalc.Constants;
   var tokentype = scf.TokenType;
   var token_precedence = scf.TokenPrecedence;

   var revpolish = [];
   var parsestack = [];

   var errortext = "";

   var function_start = -1;

   var i, pii, ttype, ttext, tprecedence, tstackprecedence;

   for (i=0; i<parseinfo.length; i++) {
      pii = parseinfo[i];
      ttype = pii.type;
      ttext = pii.text;
      if (ttype == tokentype.num || ttype == tokentype.coord || ttype == tokentype.string) {
         revpolish.push(i);
         }
      else if (ttype == tokentype.name) {
         parsestack.push(i);
         revpolish.push(function_start);
         }
      else if (ttype == tokentype.space) { // ignore
         continue;
         }
      else if (ttext == ',') {
         while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].text != "(") {
            revpolish.push(parsestack.pop());
            }
         if (parsestack.length == 0) { // no ( -- error
            errortext = scc.s_parseerrmissingopenparen;
            break;
            }
         }
      else if (ttext == '(') {
         parsestack.push(i);
         }
      else if (ttext == ')') {
         while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].text != "(") {
            revpolish.push(parsestack.pop());
            }
         if (parsestack.length == 0) { // no ( -- error
            errortext = scc.s_parseerrcloseparennoopen;
            break;
            }
         parsestack.pop();
         if (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.name) {
            revpolish.push(parsestack.pop());
            }
         }
      else if (ttype == tokentype.op) {
         if (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.name) {
            revpolish.push(parsestack.pop());
            }
         while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.op
                && parseinfo[parsestack[parsestack.length-1]].text != '(') {
            tprecedence = token_precedence[pii.opcode];
            tstackprecedence = token_precedence[parseinfo[parsestack[parsestack.length-1]].opcode];
            if (tprecedence >= 0 && tprecedence < tstackprecedence) {
               break;
               }
            else if (tprecedence < 0) {
               tprecedence = -tprecedence;
               if (tstackprecedence < 0) tstackprecedence = -tstackprecedence;
               if (tprecedence <= tstackprecedence) {
                  break;
                  }
               }
            revpolish.push(parsestack.pop());
            }
         parsestack.push(i);
         }
      else if (ttype == tokentype.error) {
         errortext = ttext;
         break;
         }
      else {
         errortext = "Internal error while processing parsed formula. ";
         break;
         }
      }
   while (parsestack.length>0) {
      if (parseinfo[parsestack[parsestack.length-1]].text == '(') {
         errortext = scc.s_parseerrmissingcloseparen;
         break;
         }
      revpolish.push(parsestack.pop());
      }

   if (errortext) {
      return errortext;
      }

   return revpolish;

   }


//
// result = SocialCalc.Formula.EvaluatePolish(parseinfo, revpolish, sheet, allowrangereturn)
//
// Execute reverse polish representation of formula
//
// Operand values are objects in the operand array with a "type" and an optional "value".
// Type can have these values (many are type and sub-type as two or more letters):
//    "tw", "th", "t", "n", "nt", "coord", "range", "start", "eErrorType", "b" (blank)
// The value of a coord is in the form A57 or A57!sheetname
// The value of a range is coord|coord|number where number starts at 0 and is
// the offset of the next item to fetch if you are going through the range one by one
// The number starts as a null string ("A1|B3|")
//

SocialCalc.Formula.EvaluatePolish = function(parseinfo, revpolish, sheet, allowrangereturn) {

   var scf = SocialCalc.Formula;
   var scc = SocialCalc.Constants;
   var tokentype = scf.TokenType;
   var lookup_result_type = scf.LookupResultType;
   var typelookup = scf.TypeLookupTable;
   var operand_as_number = scf.OperandAsNumber;
   var operand_as_text = scf.OperandAsText;
   var operand_value_and_type = scf.OperandValueAndType;
   var operands_as_coord_on_sheet = scf.OperandsAsCoordOnSheet;
   var format_number_for_display = SocialCalc.format_number_for_display || function(v, t, f) {return v+"";};

   var errortext = "";
   var function_start = -1;
   var missingOperandError = {value: "", type: "e#VALUE!", error: scc.s_parseerrmissingoperand};

   var operand = [];
   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   var i, rii, prii, ttype, ttext, value, value1, value2, tostype, tostype2, resulttype, valuetype, cond, vmatch, smatch;

   if (!parseinfo.length || (! (revpolish instanceof Array))) {
      return ({value: "", type: "e#VALUE!", error: (typeof revpolish == "string" ? revpolish : "")});
      }

   for (i=0; i<revpolish.length; i++) {
      rii = revpolish[i];
      if (rii == function_start) { // Remember the start of a function argument list
         PushOperand("start", 0);
         continue;
         }

      prii = parseinfo[rii];
      ttype = prii.type;
      ttext = prii.text;

      if (ttype == tokentype.num) {
         PushOperand("n", ttext-0);
         }

      else if (ttype == tokentype.coord) {
         PushOperand("coord", ttext);
         }

      else if (ttype == tokentype.string) {
         PushOperand("t", ttext);
         }

      else if (ttype == tokentype.op) {
         if (operand.length <= 0) { // Nothing on the stack...
            return missingOperandError;
            break; // done
            }

         // Unary minus

         if (ttext == 'M') {
            value1 = operand_as_number(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.unaryminus);
            PushOperand(resulttype, -value1.value);
            }

         // Unary plus

         else if (ttext == 'P') {
            value1 = operand_as_number(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.unaryplus);
            PushOperand(resulttype, value1.value);
            }

         // Unary % - percent, left associative

         else if (ttext == '%') {
            value1 = operand_as_number(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.unarypercent);
            PushOperand(resulttype, 0.01*value1.value);
            }

         // & - string concatenate

         else if (ttext == '&') {
            if (operand.length <= 1) { // Need at least two things on the stack...
               return missingOperandError;
               }
            value2 = operand_as_text(sheet, operand);
            value1 = operand_as_text(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.concat);
            PushOperand(resulttype, value1.value + value2.value);
            }

         // : - Range constructor

         else if (ttext == ':') {
            if (operand.length <= 1) { // Need at least two things on the stack...
               return missingOperandError;
               }
            value1 = scf.OperandsAsRangeOnSheet(sheet, operand); // get coords even if use name on other sheet
            if (value1.error) { // not available
               errortext = errortext || value1.error;
               }
            PushOperand(value1.type, value1.value); // push sheetname with range on that sheet
            }

         // ! - sheetname!coord

         else if (ttext == '!') {
            if (operand.length <= 1) { // Need at least two things on the stack...
               return missingOperandError;
               }
            value1 = operands_as_coord_on_sheet(sheet, operand); // get coord even if name on other sheet
            if (value1.error) { // not available
               errortext = errortext || value1.error;
               }
            PushOperand(value1.type, value1.value); // push sheetname with coord or range on that sheet
            }

         // Comparison operators: < L = G > N (< <= = >= > <>)

         else if (ttext == "<" || ttext == "L" || ttext == "=" || ttext == "G" || ttext == ">" || ttext == "N") {
            if (operand.length <= 1) { // Need at least two things on the stack...
               errortext = scc.s_parseerrmissingoperand; // remember error
               break;
               }
            value2 = operand_value_and_type(sheet, operand);
            value1 = operand_value_and_type(sheet, operand);
            if (value1.type.charAt(0) == "n" && value2.type.charAt(0) == "n") { // compare two numbers
               cond = 0;
               if (ttext == "<") { cond = value1.value < value2.value ? 1 : 0; }
               else if (ttext == "L") { cond = value1.value <= value2.value ? 1 : 0; }
               else if (ttext == "=") { cond = value1.value == value2.value ? 1 : 0; }
               else if (ttext == "G") { cond = value1.value >= value2.value ? 1 : 0; }
               else if (ttext == ">") { cond = value1.value > value2.value ? 1 : 0; }
               else if (ttext == "N") { cond = value1.value != value2.value ? 1 : 0; }
               PushOperand("nl", cond);
               }
            else if (value1.type.charAt(0) == "e") { // error on left
               PushOperand(value1.type, 0);
               }               
            else if (value2.type.charAt(0) == "e") { // error on right
               PushOperand(value2.type, 0);
               }               
            else { // text maybe mixed with numbers or blank
               tostype = value1.type.charAt(0);
               tostype2 = value2.type.charAt(0);
               if (tostype == "n") {
                  value1.value = format_number_for_display(value1.value, "n", "");
                  }
               else if (tostype == "b") {
                  value1.value = "";
                  }
               if (tostype2 == "n") {
                  value2.value = format_number_for_display(value2.value, "n", "");
                  }
               else if (tostype2 == "b") {
                  value2.value = "";
                  }
               cond = 0;
               value1.value = value1.value.toLowerCase(); // ignore case
               value2.value = value2.value.toLowerCase();
               if (ttext == "<") { cond = value1.value < value2.value ? 1 : 0; }
               else if (ttext == "L") { cond = value1.value <= value2.value ? 1 : 0; }
               else if (ttext == "=") { cond = value1.value == value2.value ? 1 : 0; }
               else if (ttext == "G") { cond = value1.value >= value2.value ? 1 : 0; }
               else if (ttext == ">") { cond = value1.value > value2.value ? 1 : 0; }
               else if (ttext == "N") { cond = value1.value != value2.value ? 1 : 0; }
               PushOperand("nl", cond);
               }
            }

         // Normal infix arithmethic operators: +, -. *, /, ^

         else { // what's left are the normal infix arithmetic operators
            if (operand.length <= 1) { // Need at least two things on the stack...
               errortext = scc.s_parseerrmissingoperand; // remember error
               break;
               }
            value2 = operand_as_number(sheet, operand);
            value1 = operand_as_number(sheet, operand);
            if (ttext == '+') {
               resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
               PushOperand(resulttype, value1.value + value2.value);
               }
            else if (ttext == '-') {
               resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
               PushOperand(resulttype, value1.value - value2.value);
               }
            else if (ttext == '*') {
               resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
               PushOperand(resulttype, value1.value * value2.value);
               }
            else if (ttext == '/') {
               if (value2.value != 0) {
                  PushOperand("n", value1.value / value2.value); // gives plain numeric result type
                  }
               else {
                  PushOperand("e#DIV/0!", 0);
                  }
               }
            else if (ttext == '^') {
               value1.value = Math.pow(value1.value, value2.value);
               value1.type = "n"; // gives plain numeric result type
               if (isNaN(value1.value)) {
                  value1.value = 0;
                  value1.type = "e#NUM!";
                  }
               PushOperand(value1.type, value1.value);
               }
            }
         }

      // function or name

      else if (ttype == tokentype.name) {
         errortext = scf.CalculateFunction(ttext, operand, sheet);
         if (errortext) break;
         }

      else {
         errortext = scc.s_InternalError+"Unknown token "+ttype+" ("+ttext+"). ";
         break;
         }
      }

   // look at final value and handle special cases

   value = operand[0] ? operand[0].value : "";
   tostype = operand[0] ? operand[0].type : "";

   if (tostype == "name") { // name - expand it
      value1 = SocialCalc.Formula.LookupName(sheet, value);
      value = value1.value;
      tostype = value1.type;
      errortext = errortext || value1.error;
      }

   if (tostype == "coord") { // the value is a coord reference, get its value and type
      value1 = operand_value_and_type(sheet, operand);
      value = value1.value;
      tostype = value1.type;
      if (tostype == "b") {
         tostype = "n";
         value = 0;
         }
      }

   if (operand.length > 1 && !errortext) { // something left - error
      errortext += scc.s_parseerrerrorinformula;
      }

   // set return type

   valuetype = tostype;

   if (tostype.charAt(0) == "e") { // error value
      errortext = errortext || tostype.substring(1) || scc.s_calcerrerrorvalueinformula;
      }
   else if (tostype == "range") {
      vmatch = value.match(/^(.*)\|(.*)\|/);
      smatch = vmatch[1].indexOf("!");
      if (smatch>=0) { // swap sheetname
         vmatch[1] = vmatch[1].substring(smatch+1) + "!" + vmatch[1].substring(0, smatch).toUpperCase();
         }
      else {
         vmatch[1] = vmatch[1].toUpperCase();
         }
      value = vmatch[1] + ":" + vmatch[2].toUpperCase();
      if (!allowrangereturn) {
         errortext = scc.s_formularangeresult+" "+value;
         }
      }

   if (errortext && valuetype.charAt(0) != "e") {
      value = errortext;
      valuetype = "e";
     }

   // look for overflow

   if (valuetype.charAt(0) == "n" && (isNaN(value) || !isFinite(value))) {
      value = 0;
      valuetype = "e#NUM!";
      errortext = isNaN(value) ? scc.s_calcerrnumericnan: scc.s_calcerrnumericoverflow;
      }

   return ({value: value, type: valuetype, error: errortext});

   }


/*
#
# resulttype = SocialCalc.Formula.LookupResultType(type1, type2, typelookup);
#
# typelookup has values of the following form:
#
#    typelookup{"typespec1"} = "|typespec2A:resultA|typespec2B:resultB|..."
#
# First type1 is looked up. If no match, then the first letter (major type) of type1 plus "*" is looked up
# resulttype is type1 if result is "1", type2 if result is "2", otherwise the value of result.
#
*/

SocialCalc.Formula.LookupResultType = function(type1, type2, typelookup) {

   var pos1, pos2, result;

   var table1 = typelookup[type1];

   if (!table1) {
      table1 = typelookup[type1.charAt(0)+'*'];
      if (!table1) {
         return "e#VALUE! (internal error, missing LookupResultType "+type1.charAt(0)+"*)"; // missing from table -- please add it
         }
      }
   pos1 = table1.indexOf("|"+type2+":");
   if (pos1 >= 0) {
      pos2 = table1.indexOf("|", pos1+1);
      if (pos2<0) return "e#VALUE! (internal error, incorrect LookupResultType "+table1+")";
      result = table1.substring(pos1+type2.length+2, pos2);
      if (result == "1") return type1;
      if (result == "2") return type2;
      return result;
      }
   pos1 = table1.indexOf("|"+type2.charAt(0)+"*:");
   if (pos1 >= 0) {
      pos2 = table1.indexOf("|", pos1+1);
      if (pos2<0) return "e#VALUE! (internal error, incorrect LookupResultType "+table1+")";
      result = table1.substring(pos1+4, pos2);
      if (result == "1") return type1;
      if (result == "2") return type2;
      return result;
      }
   return "e#VALUE!";

   }

/*
#
# operandinfo = SocialCalc.Formula.TopOfStackValueAndType(sheet, operand)
#
# Returns top of stack value and type and then pops the stack.
# The result is {value: value, type: type, error: "only if bad error"}
#
*/

SocialCalc.Formula.TopOfStackValueAndType = function(sheet, operand) {

   var cellvtype, cell, pos, coordsheet;
   var scf = SocialCalc.Formula;

   var result = {type: "", value: ""};

   var stacklen = operand.length;

   if (!stacklen) { // make sure something is there
      result.error = SocialCalc.Constants.s_InternalError+"no operand on stack";
      return result;
      }

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   if (result.type == "name") {
      result = scf.LookupName(sheet, result.value);
      }

   return result;

   }


/*
#
# operandinfo = OperandAsNumber(sheet, operand)
#
# Uses operand_value_and_type to get top of stack and pops it.
# Returns numeric value and type.
# Text values are treated as 0 if they can't be converted somehow.
#
*/

SocialCalc.Formula.OperandAsNumber = function(sheet, operand) {

   var t, valueinfo;
   var operandinfo = SocialCalc.Formula.OperandValueAndType(sheet, operand);

   t = operandinfo.type.charAt(0);

   if (t == "n") {
      operandinfo.value = operandinfo.value-0;
      }
   else if (t == "b") { // blank cell
      operandinfo.type = "n";
      operandinfo.value = 0;
      }
   else if (t == "e") { // error
      operandinfo.value = 0;
      }
   else {
      valueinfo = SocialCalc.DetermineValueType ? SocialCalc.DetermineValueType(operandinfo.value) :
                                                    {value: operandinfo.value-0, type: "n"}; // if without rest of SocialCalc
      if (valueinfo.type.charAt(0) == "n") {
         operandinfo.value = valueinfo.value-0;
         operandinfo.type = valueinfo.type;
         }
      else {
         operandinfo.value = 0;
         operandinfo.type = valueinfo.type;
         }
      }

   return operandinfo;

   }

/*
#
# operandinfo = OperandAsText(sheet, operand)
#
# Uses operand_value_and_type to get top of stack and pops it.
# Returns text value, preserving sub-type.
#
*/

SocialCalc.Formula.OperandAsText = function(sheet, operand) {

   var t, valueinfo;
   var operandinfo = SocialCalc.Formula.OperandValueAndType(sheet, operand);

   t = operandinfo.type.charAt(0);

   if (t ==  "t") { // any flavor of text returns as is
      ;
      }
   else if (t == "n") {
      operandinfo.value = SocialCalc.format_number_for_display ?
                             SocialCalc.format_number_for_display(operandinfo.value, operandinfo.type, "") :
                             operandinfo.value = operandinfo.value+"";
      operandinfo.type = "t";
      }
   else if (t == "b") { // blank
      operandinfo.value = "";
      operandinfo.type = "t";
      }
   else if (t == "e") { // error
      operandinfo.value = "";
      }
   else {
      operand.value = operandinfo.value + "";
      operand.type = "t";
      }

   return operandinfo;

   }

/*
#
# result = SocialCalc.Formula.OperandValueAndType(sheet, operand)
#
# Pops the top of stack and returns it, following a coord reference if necessary.
# The result is {value: value, type: type, error: "only if bad error"}
# Ranges are returned as if they were pushed onto the stack first coord first
# Also sets type with "t", "n", "th", etc., as appropriate
#
*/

SocialCalc.Formula.OperandValueAndType = function(sheet, operand) {

   var cellvtype, cell, pos, coordsheet;
   var scf = SocialCalc.Formula;

   var result = {type: "", value: ""};

   var stacklen = operand.length;

   if (!stacklen) { // make sure something is there
      result.error = SocialCalc.Constants.s_InternalError+"no operand on stack";
      return result;
      }

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   if (result.type == "name") {
      result = scf.LookupName(sheet, result.value);
      }

   if (result.type == "range") {
      result = scf.StepThroughRangeDown(operand, result.value);
      }

   if (result.type == "coord") { // value is a coord reference
      coordsheet = sheet;
      pos = result.value.indexOf("!");
      if (pos != -1) { // sheet reference
         coordsheet = scf.FindInSheetCache(result.value.substring(pos+1)); // get other sheet
         if (coordsheet == null) { // unavailable
            result.type = "e#REF!";
            result.error = SocialCalc.Constants.s_sheetunavailable+" "+result.value.substring(pos+1);
            result.value = 0;
            return result;
            }
         result.value = result.value.substring(0, pos); // get coord part
         }

      if (coordsheet) {
         cell = coordsheet.cells[SocialCalc.Formula.PlainCoord(result.value)];
         if (cell) {
            cellvtype = cell.valuetype; // get type of value in the cell it points to
            result.value = cell.datavalue;
            }
         else {
            cellvtype = "b";
            }
         }
      else {
         cellvtype = "e#N/A";
         result.value = 0;
         }
      result.type = cellvtype || "b";
      if (result.type == "b") { // blank
         result.value = 0;
         }
      }

   return result;

   }

/*
#
# operandinfo = SocialCalc.Formula.OperandAsCoord(sheet, operand)
#
# Gets top of stack and pops it.
# Returns coord value. All others are treated as an error.
#
*/


SocialCalc.Formula.OperandAsCoord = function(sheet, operand) {

   var scf = SocialCalc.Formula;

   var result = {type: "", value: ""};

   var stacklen = operand.length;

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack
   if (result.type == "name") {
      result = SocialCalc.Formula.LookupName(sheet, result.value);
      }
   if (result.type == "coord") { // value is a coord reference
      return result;
      }
   else {
      result.value = SocialCalc.Constants.s_calcerrcellrefmissing;
      result.type = "e#REF!";
      return result;
      }
}


/*
#
# result = SocialCalc.Formula.OperandsAsCoordOnSheet(sheet, operand)
#
# Gets 2 at top of stack and pops them, treating them as sheetname!coord-or-name.
# Returns stack-style coord value (coord!sheetname, or coord!sheetname|coord|) with
# a type of coord or range. All others are treated as an error.
# If sheetname not available, sets result.error.
#
*/

SocialCalc.Formula.OperandsAsCoordOnSheet = function(sheet, operand) {

   var sheetname, othersheet, pos1, pos2;
   var value1 = {};
   var result = {};
   var scf = SocialCalc.Formula;

   var stacklen = operand.length;
   value1.value = operand[stacklen-1].value; // get top of stack - coord or name
   value1.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   sheetname = scf.OperandAsSheetName(sheet, operand); // get sheetname as text
   othersheet = scf.FindInSheetCache(sheetname.value);
   if (othersheet == null) { // unavailable
      result.type = "e#REF!";
      result.value = 0;
      result.error = SocialCalc.Constants.s_sheetunavailable+" "+sheetname.value;
      return result;
      }

   if (value1.type == "name") {
      value1 = scf.LookupName(othersheet, value1.value);
      }
   result.type = value1.type;
   if (value1.type == "coord") { // value is a coord reference
      result.value = value1.value + "!" + sheetname.value; // return in the format as used on stack
      }
   else if (value1.type == "range") { // value is a range reference
      pos1 = value1.value.indexOf("|");
      pos2 = value1.value.indexOf("|", pos1+1);
      result.value = value1.value.substring(0, pos1) + "!" + sheetname.value +
                    "|" + value1.value.substring(pos1+1, pos2) + "|";
      }
   else if (value1.type.charAt(0)=="e") {
      result.value = value1.value;
      }
   else {
      result.error = SocialCalc.Constants.s_calcerrcellrefmissing;
      result.type = "e#REF!";
      result.value = 0;
      }
   return result;
   
   }

/*
#
# result = SocialCalc.Formula.OperandsAsRangeOnSheet(sheet, operand)
#
# Gets 2 at top of stack and pops them, treating them as coord2-or-name:coord1.
# Name is evaluated on sheet of coord1.
# Returns result with "value" of stack-style range value (coord!sheetname|coord|) and
# "type" of "range". All others are treated as an error.
#
*/

SocialCalc.Formula.OperandsAsRangeOnSheet = function(sheet, operand) {

   var value1, othersheet, pos1, pos2;
   var value2 = {};
   var scf = SocialCalc.Formula;
   var scc = SocialCalc.Constants;

   var stacklen = operand.length;
   value2.value = operand[stacklen-1].value; // get top of stack - coord or name for "right" side
   value2.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   value1 = scf.OperandAsCoord(sheet, operand); // get "left" coord
   if (value1.type != "coord") { // not a coord, which it must be
      return {value: 0, type: "e#REF!"};
      }

   othersheet = sheet;
   pos1 = value1.value.indexOf("!");
   if (pos1 != -1) { // sheet reference
      pos2 = value1.value.indexOf("|", pos1+1);
      if (pos2 < 0) pos2 = value1.value.length;
      othersheet = scf.FindInSheetCache(value1.value.substring(pos1+1,pos2)); // get other sheet
      if (othersheet == null) { // unavailable
         return {value: 0, type: "e#REF!", errortext: scc.s_sheetunavailable+" "+value1.value.substring(pos1+1,pos2)};
         }
      }

   if (value2.type == "name") { // coord:name is allowed, if name is just one cell
      value2 = scf.LookupName(othersheet, value2.value, "end");
      }

   if (value2.type == "coord") { // value is a coord reference, so return the combined range
      return {value: value1.value+"|"+value2.value+"|", type: "range"}; // return range in the format as used on stack
      }
   else { // bad form
      return {value: scc.s_calcerrcellrefmissing, type: "e#REF!"};
      }
   }


/*
#
# result = SocialCalc.Formula.OperandAsSheetName(sheet, operand)
#
# Gets top of stack and pops it.
# Returns sheetname value. All others are treated as an error.
# Accepts text, cell reference, and named value which is one of those two.
#
*/

SocialCalc.Formula.OperandAsSheetName = function(sheet, operand) {

   var nvalue, cell;

   var scf = SocialCalc.Formula;

   var result = {type: "", value: ""};

   var stacklen = operand.length;

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack
   if (result.type == "name") {
      nvalue = SocialCalc.Formula.LookupName(sheet, result.value);
      if (!nvalue.value) { // not a known name - return bare name as the name value
         return result;
         }
      result.value = nvalue.value;
      result.type = nvalue.type;
      }
   if (result.type == "coord") { // value is a coord reference, follow it to find sheet name
      cell = sheet.cells[SocialCalc.Formula.PlainCoord(result.value)];
      if (cell) {
         result.value = cell.datavalue;
         result.type = cell.valuetype;
         }
      else {
         result.value = "";
         result.type = "b";
         }
      }
   if (result.type.charAt(0) == "t") { // value is a string which could be a sheet name
      return result;
      }
   else {
      result.value = "";
      result.error = SocialCalc.Constants.s_calcerrsheetnamemissing;
      return result;
      }

   }

//
// value = SocialCalc.Formula.LookupName(sheet, name)
//
// Returns value and type of a named value
// Names are case insensitive
// Names may have a definition which is a coord (A1), a range (A1:B7), or a formula (=OFFSET(A1,0,0,5,1))
// Note: The range must not have sheet names ("!") in them.
//

SocialCalc.Formula.LookupName = function(sheet, name, isEnd) {

   var pos, specialc, parseinfo;
   var names = sheet.names;
   var value = {};
   var startedwalk = false;

   if (names[name.toUpperCase()]) { // is name defined?

      value.value = names[name.toUpperCase()].definition; // yes

      if (value.value.charAt(0) == "=") { // formula
         if (!sheet.checknamecirc) { // are we possibly walking the name tree?
            sheet.checknamecirc = {}; // not yet
            startedwalk = true; // remember we are the reference that started it
            }
         else {
            if (sheet.checknamecirc[name]) { // circular reference
               value.type = "e#NAME?";
               value.error = SocialCalc.Constants.s_circularnameref+' "' + name + '".';
               return value;
               }
            }
         sheet.checknamecirc[name] = true;

         parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(value.value.substring(1));
         value = SocialCalc.Formula.evaluate_parsed_formula(parseinfo, sheet, 1); // parse formula, allowing range return

         delete sheet.checknamecirc[name]; // done with us
         if (startedwalk) {
            delete sheet.checknamecirc; // done with walk
            }

         if (value.type != "range") {
            return value;
            }
         }

      pos = value.value.indexOf(":");
      if (pos != -1) { // range
         value.type = "range";
         value.value = value.value.substring(0, pos) + "|" + value.value.substring(pos+1)+"|";
         value.value = value.value.toUpperCase();
         }
      else {
         value.type = "coord";
         value.value = value.value.toUpperCase();
         }
      return value;
      }
   else if (specialc=SocialCalc.Formula.SpecialConstants[name.toUpperCase()]) { // special constant, like #REF!
      pos = specialc.indexOf(",");
      value.value = specialc.substring(0,pos)-0;
      value.type = specialc.substring(pos+1);
      return value;
      }
   else if (/^[a-zA-Z][a-zA-Z]?$/.test(name)) {
      value.type = "coord";
      value.value = name.toUpperCase() + (isEnd ? sheet.attribs.lastrow : 1);
      return value;
   }
   else {
      value.value = "";
      value.type = "e#NAME?";
      value.error = SocialCalc.Constants.s_calcerrunknownname+' "'+name+'"';
      return value;
      }
   }

/*
#
# coord = SocialCalc.Formula.StepThroughRangeDown(operand, rangevalue)
#
# Returns next coord in a range, keeping track on the operand stack
# Goes from upper left across and down to bottom right.
#
*/

SocialCalc.Formula.StepThroughRangeDown = function(operand, rangevalue) {

   var value1, value2, sequence, pos1, pos2, sheet1, rp, c, r, count;
   var scf = SocialCalc.Formula;

   pos1 = rangevalue.indexOf("|");
   pos2 = rangevalue.indexOf("|", pos1+1);
   value1 = rangevalue.substring(0, pos1);
   value2 = rangevalue.substring(pos1+1, pos2);
   sequence = rangevalue.substring(pos2+1) - 0;

   pos1 = value1.indexOf("!");
   if (pos1 != -1) {
      sheet1 = value1.substring(pos1);
      value1 = value1.substring(0, pos1);
      }
   else {
      sheet1 = "";
      }
   pos1 = value2.indexOf("!");
   if (pos1 != -1) {
      value2 = value2.substring(0, pos1);
      }

   rp = scf.OrderRangeParts(value1, value2);
   
   count = 0;
   for (r=rp.r1; r<=rp.r2; r++) {
      for (c=rp.c1; c<=rp.c2; c++) {
         count++;
         if (count > sequence) {
            if (r!=rp.r2 || c!=rp.c2) { // keep on stack until done
               scf.PushOperand(operand, "range", value1+sheet1+"|"+value2+"|"+count);
               }
            return {value: SocialCalc.crToCoord(c, r)+sheet1, type: "coord"};
            }
         }
      }
   }

/*
#
# result = SocialCalc.Formula.DecodeRangeParts(sheetdata, range)
#
# Returns sheetdata for the sheet where the range is, as well as
# the number of the first column in the range, the number of columns,
# and equivalent row information:
#
# {sheetdata: sheet, sheetname: name-or-"", col1num: n, ncols: n, row1num: n, nrows: n}
#
# If any errors, a null result is returned.
#
*/

SocialCalc.Formula.DecodeRangeParts = function(sheetdata, range) {

   var value1, value2, pos1, pos2, sheet1, coordsheetdata, rp;

   var scf = SocialCalc.Formula;

   pos1 = range.indexOf("|");
   pos2 = range.indexOf("|", pos1+1);
   value1 = range.substring(0, pos1);
   value2 = range.substring(pos1+1, pos2);

   pos1 = value1.indexOf("!");
   if (pos1 != -1) {
      sheet1 = value1.substring(pos1+1);
      value1 = value1.substring(0, pos1);
      }
   else {
      sheet1 = "";
      }
   pos1 = value2.indexOf("!");
   if (pos1 != -1) {
      value2 = value2.substring(0, pos1);
      }

   coordsheetdata = sheetdata;
   if (sheet1) { // sheet reference
      coordsheetdata = scf.FindInSheetCache(sheet1);
      if (coordsheetdata == null) { // unavailable
         return null;
         }
      }

   rp = scf.OrderRangeParts(value1, value2);

   return {sheetdata: coordsheetdata, sheetname: sheet1, col1num: rp.c1, ncols: rp.c2-rp.c1+1, row1num: rp.r1, nrows: rp.r2-rp.r1+1}

   }


//*********************
//
// Function Handling
//
//*********************

// List of functions -- Define after functions are defined
//
// SocialCalc.Formula.FunctionList["function_name"] = [function_subroutine, number_of_arguments, arg_def, func_def, func_class]
//   function_subroutine takes arguments (fname, operand, foperand, sheet), returns
//      errortext or null, pushing result on operand stack.
//   number_of_arguments is:
//      0 = no arguments
//      >0 = exactly that many arguments
//      <0 = that many arguments (abs value) or more
//      100 = don't check
//
//   arg_def, if present, is the name of the element in SocialCalc.Formula.FunctionArgDefs.
//   func_def, if present, is a string explaining the function. If not, looked up in SocialCalc.Constants.
//   func_class, if present, is the comma-separated names of the elements in SocialCalc.Formula.FunctionClasses.
//
// To add a function, just add it to this object.

   if (!SocialCalc.Formula.FunctionList) { // make sure it is defined (could have been in another module)
      SocialCalc.Formula.FunctionList = {};
      }

   // FunctionClasses[classname] = {name: full-name-string, items: [sorted list of function names]};
   // filled in by SocialCalc.Formula.FillFunctionInfo

   SocialCalc.Formula.FunctionClasses = null; // start null to say needs filling in

   // FunctionArgDef[argname] = explicit-string-for-arg-list;
   // filled in by SocialCalc.Formula.FillFunctionInfo

   SocialCalc.Formula.FunctionArgDefs = {};


/*
#
# errortext = SocialCalc.Formula.CalculateFunction(fname, operand, sheet)
#
# Dispatches for function fname.
#
*/

SocialCalc.Formula.CalculateFunction = function(fname, operand, sheet) {

   var fobj, foperand, ffunc, argnum, ttext;
   var scf = SocialCalc.Formula;
   var ok = 1;
   var errortext = "";

   fobj = scf.FunctionList[fname];

   if (fobj) {
      foperand = [];
      ffunc = fobj[0];
      argnum = fobj[1];
      scf.CopyFunctionArgs(operand, foperand);
      if (argnum != 100) {
         if (argnum < 0) {
            if (foperand.length < -argnum) {
               errortext = scf.FunctionArgsError(fname, operand);
               return errortext;
               }
            }
         else {
            if (foperand.length != argnum) {
               errortext = scf.FunctionArgsError(fname, operand);
               return errortext;
               }
            }
         }
      errortext = ffunc(fname, operand, foperand, sheet);
      }

   else {
         ttext = fname;

         if (operand.length && operand[operand.length-1].type == "start") { // no arguments - name or zero arg function
            operand.pop();
            scf.PushOperand(operand, "name", ttext);
            }

         else {
            errortext = SocialCalc.Constants.s_sheetfuncunknownfunction+" " + ttext +". ";
            }
      }

   return errortext;

}

//
// SocialCalc.Formula.PushOperand(operand, t, v)
//
// Pushes the type and value onto the operand stack
//

SocialCalc.Formula.PushOperand = function(operand, t, v) {

   operand.push({type: t, value: v});

   }

//
// SocialCalc.Formula.CopyFunctionArgs(operand, foperand)
//
// Pops operands from operand and pushes on foperand up to function start
// reversing order in the process.
//

SocialCalc.Formula.CopyFunctionArgs = function(operand, foperand) {

   var fobj, foperand, ffunc, argnum;
   var scf = SocialCalc.Formula;
   var ok = 1;
   var errortext = null;

   while (operand.length>0 && operand[operand.length-1].type != "start") { // get each arg
      foperand.push(operand.pop()); // copy it
      }
   operand.pop(); // get rid of "start"

   return;

   }

//
// errortext = SocialCalc.Formula.FunctionArgsError(fname, operand)
//
// Pushes appropriate error on operand stack and returns errortext, including fname
//

SocialCalc.Formula.FunctionArgsError = function(fname, operand) {

   var errortext = SocialCalc.Constants.s_calcerrincorrectargstofunction+" " + fname + ". ";
   SocialCalc.Formula.PushOperand(operand, "e#VALUE!", errortext);

   return errortext;

   }


//
// errortext = SocialCalc.Formula.FunctionSpecificError(fname, operand, errortype, errortext)
//
// Pushes specified error and text on operand stack.
//

SocialCalc.Formula.FunctionSpecificError = function(fname, operand, errortype, errortext) {

   SocialCalc.Formula.PushOperand(operand, errortype, errortext);

   return errortext;

   }

//
// haserror = SocialCalc.Formula.CheckForErrorValue(operand, v)
//
// If v.type is an error, push it on operand stack and return true, otherwise return false.
//

SocialCalc.Formula.CheckForErrorValue = function(operand, v) {

   if (v.type.charAt(0) == "e") {
      operand.push(v);
      return true;
      }
   else {
      return false;
      }

   }

/////////////////////////
//
// FUNCTION INFORMATION ROUTINES
//

//
// SocialCalc.Formula.FillFunctionInfo()
//
// Goes through function definitions and fills out FunctionArgDefs and FunctionClasses.
// Execute this after any changes to SocialCalc.Constants but before UI is used.
//

SocialCalc.Formula.FillFunctionInfo = function() {

   var scf = SocialCalc.Formula;
   var scc = SocialCalc.Constants;

   var fname, f, classes, cname, i;

   if (scf.FunctionClasses) { // only do once
      return;
      }

   for (fname in scf.FunctionList) {
      f = scf.FunctionList[fname];
      if (f[2]) { // has an arg def
         scf.FunctionArgDefs[f[2]] = scc["s_farg_"+f[2]] || ""; // get it from constants
         }
      if (!f[3]) { // no text def, see if in constants
         if (scc["s_fdef_"+fname]) {
            scf.FunctionList[fname][3] = scc["s_fdef_"+fname];
            }
         }
      }

   scf.FunctionClasses = {};
 
   for (i=0; i<scc.function_classlist.length; i++) {
      cname = scc.function_classlist[i];
      scf.FunctionClasses[cname] = {name: scc["s_fclass_"+cname], items: []};
      }

   for (fname in scf.FunctionList) {
      f = scf.FunctionList[fname];
      classes = f[4] ? f[4].split(",") : []; // get classes
      classes.push("all");
      for (i=0; i<classes.length; i++) {
         cname = classes[i];
         scf.FunctionClasses[cname].items.push(fname);
         }
      }
   for (cname in scf.FunctionClasses) {
      scf.FunctionClasses[cname].items.sort();
      }

   }

//
// str = SocialCalc.Formula.FunctionArgString(fname)
//
// Returns a string representing the arguments to function fname.
//

SocialCalc.Formula.FunctionArgString = function(fname) {

   var scf = SocialCalc.Formula;
   var fdata = scf.FunctionList[fname];
   var nargs, i, str;

   var adef = fdata[2];

   if (!adef) {
      nargs = fdata[1];
      if (nargs == 0) {
         adef = " ";
         }
      else if (nargs > 0) {
         str = "v1";
         for (i=2; i<=nargs; i++) {
            str += ", v"+i;
            }
         return str;
         }
      else if (nargs < 0) {
         str = "v1";
         for (i=2; i<-nargs; i++) {
            str += ", v"+i;
            }
         return str+", ...";
         }
      else {
         return "nargs: "+nargs;
         }
      }

   str = scf.FunctionArgDefs[adef] || adef;

   return str;

   }


/////////////////////////
//
// FUNCTION DEFINITIONS
//
// The standard function definitions follow.
//
// Note that some need SocialCalc.DetermineValueType to be defined.
//

/*
#
# AVERAGE(v1,c1:c2,...)
# COUNT(v1,c1:c2,...)
# COUNTA(v1,c1:c2,...)
# COUNTBLANK(v1,c1:c2,...)
# MAX(v1,c1:c2,...)
# MIN(v1,c1:c2,...)
# PRODUCT(v1,c1:c2,...)
# STDEV(v1,c1:c2,...)
# STDEVP(v1,c1:c2,...)
# SUM(v1,c1:c2,...)
# VAR(v1,c1:c2,...)
# VARP(v1,c1:c2,...)
#
# Calculate all of these and then return the desired one (overhead is in accessing not calculating)
# If this routine is changed, check the dseries_functions, too.
#
*/

SocialCalc.Formula.SeriesFunctions = function(fname, operand, foperand, sheet) {

   var value1, t, v1;

   var scf = SocialCalc.Formula;
   var operand_value_and_type = scf.OperandValueAndType;
   var lookup_result_type = scf.LookupResultType;
   var typelookupplus = scf.TypeLookupTable.plus;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   var sum = 0;
   var resulttypesum = "";
   var count = 0;
   var counta = 0;
   var countblank = 0;
   var product = 1;
   var maxval;
   var minval;
   var mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                         // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

   while (foperand.length > 0) {
      value1 = operand_value_and_type(sheet, foperand);
      t = value1.type.charAt(0);
      if (t == "n") count += 1;
      if (t != "b") counta += 1;
      if (t == "b") countblank += 1;

      if (t == "n") {
         v1 = value1.value-0; // get it as a number
         sum += v1;
         product *= v1;
         maxval = (maxval!=undefined) ? (v1 > maxval ? v1 : maxval) : v1;
         minval = (minval!=undefined) ? (v1 < minval ? v1 : minval) : v1;
         if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
            mk1 = v1;
            sk1 = 0;
            }
         else { // Accumulate S sub 1 through n as per Knuth noted above
            mk = mk1 + (v1 - mk1) / count;
            sk = sk1 + (v1 - mk1) * (v1 - mk);
            sk1 = sk;
            mk1 = mk;
            }
         resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
         }
      else if (t == "e" && resulttypesum.charAt(0) != "e") {
         resulttypesum = value1.type;
         }
      }

   resulttypesum = resulttypesum || "n";

   switch (fname) {
      case "SUM":
         PushOperand(resulttypesum, sum);
         break;

      case "PRODUCT": // may handle cases with text differently than some other spreadsheets
         PushOperand(resulttypesum, product);
         break;

      case "MIN":
         PushOperand(resulttypesum, minval || 0);
         break;

      case "MAX":
         PushOperand(resulttypesum, maxval || 0);
         break;

      case "COUNT":
         PushOperand("n", count);
         break;

      case "COUNTA":
         PushOperand("n", counta);
         break;

      case "COUNTBLANK":
         PushOperand("n", countblank);
         break;

      case "AVERAGE":
         if (count > 0) {
            PushOperand(resulttypesum, sum/count);
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "STDEV":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "STDEVP":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / count));
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "VAR":
         if (count > 1) {
            PushOperand(resulttypesum, sk / (count - 1));
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "VARP":
         if (count > 1) {
            PushOperand(resulttypesum, sk / count);
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;
      }

   return null;

   }

// Add to function list
SocialCalc.Formula.FunctionList["AVERAGE"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["COUNT"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["COUNTA"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["COUNTBLANK"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["MAX"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["MIN"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["PRODUCT"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["STDEV"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["STDEVP"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["SUM"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["VAR"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["VARP"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];

/*
#
# SUMPRODUCT(range1, range2, ...)
#
*/

SocialCalc.Formula.SumProductFunction = function(fname, operand, foperand, sheet) {
  
   var range, products = [], sum = 0;
   var scf = SocialCalc.Formula;
   var ncols = 0, nrows = 0;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   while (foperand.length > 0) {
      range = scf.TopOfStackValueAndType(sheet, foperand);
      if (range.type != "range") {
         PushOperand("e#VALUE!", 0);
         return;
         }
      rangeinfo = scf.DecodeRangeParts(sheet, range.value);
      if (!ncols) ncols = rangeinfo.ncols;
      else if (ncols != rangeinfo.ncols) {
         PushOperand("e#VALUE!", 0);
         return;
         }
      if (!nrows) nrows = rangeinfo.nrows;
      else if (nrows != rangeinfo.nrows) {
         PushOperand("e#VALUE!", 0);
         return;
         }
      for (i=0; i<rangeinfo.ncols; i++) {
         for (j=0; j<rangeinfo.nrows; j++) {
            k = i * rangeinfo.nrows + j;
            cellcr = SocialCalc.crToCoord(rangeinfo.col1num + i, rangeinfo.row1num + j);
            cell = rangeinfo.sheetdata.GetAssuredCell(cellcr);
            value = cell.valuetype == "n" ? cell.datavalue : 0;
            products[k] = (products[k] || 1) * value;
            }
         }
      }
   for (i=0; i<products.length; i++) {
      sum += products[i];
      }
   PushOperand("n", sum);

   return;

   }

SocialCalc.Formula.FunctionList["SUMPRODUCT"] = [SocialCalc.Formula.SumProductFunction, -1, "rangen", "", "stat"];

/*
#
# DAVERAGE(databaserange, fieldname, criteriarange)
# DCOUNT(databaserange, fieldname, criteriarange)
# DCOUNTA(databaserange, fieldname, criteriarange)
# DGET(databaserange, fieldname, criteriarange)
# DMAX(databaserange, fieldname, criteriarange)
# DMIN(databaserange, fieldname, criteriarange)
# DPRODUCT(databaserange, fieldname, criteriarange)
# DSTDEV(databaserange, fieldname, criteriarange)
# DSTDEVP(databaserange, fieldname, criteriarange)
# DSUM(databaserange, fieldname, criteriarange)
# DVAR(databaserange, fieldname, criteriarange)
# DVARP(databaserange, fieldname, criteriarange)
#
# Calculate all of these and then return the desired one (overhead is in accessing not calculating)
# If this routine is changed, check the series_functions, too.
#
*/

SocialCalc.Formula.DSeriesFunctions = function(fname, operand, foperand, sheet) {

   var value1, tostype, cr, dbrange, fieldname, criteriarange, dbinfo, criteriainfo;
   var fieldasnum, targetcol, i, j, k, cell, criteriafieldnums;
   var testok, criteriacr, criteria, testcol, testcr;
   var t;

   var scf = SocialCalc.Formula;
   var operand_value_and_type = scf.OperandValueAndType;
   var lookup_result_type = scf.LookupResultType;
   var typelookupplus = scf.TypeLookupTable.plus;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   var value1 = {};

   var sum = 0;
   var resulttypesum = "";
   var count = 0;
   var counta = 0;
   var countblank = 0;
   var product = 1;
   var maxval;
   var minval;
   var mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                         // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

   dbrange = scf.TopOfStackValueAndType(sheet, foperand); // get a range
   fieldname = scf.OperandValueAndType(sheet, foperand); // get a value
   criteriarange = scf.TopOfStackValueAndType(sheet, foperand); // get a range

   if (dbrange.type != "range" || criteriarange.type != "range") {
      return scf.FunctionArgsError(fname, operand);
      }

   dbinfo = scf.DecodeRangeParts(sheet, dbrange.value);
   criteriainfo = scf.DecodeRangeParts(sheet, criteriarange.value);

   fieldasnum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, fieldname.value, fieldname.type);
   if (fieldasnum <= 0) {
      PushOperand("e#VALUE!", 0);
      return;
      }

   targetcol = dbinfo.col1num + fieldasnum - 1;
   criteriafieldnums = [];

   for (i=0; i<criteriainfo.ncols; i++) { // get criteria field colnums
      cell = criteriainfo.sheetdata.GetAssuredCell(SocialCalc.crToCoord(criteriainfo.col1num + i, criteriainfo.row1num));
      criterianum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, cell.datavalue, cell.valuetype);
      if (criterianum <= 0) {
         PushOperand("e#VALUE!", 0);
         return;
         }
      criteriafieldnums.push(dbinfo.col1num + criterianum - 1);
      }

   for (i=1; i<dbinfo.nrows; i++) { // go through each row of the database
      testok = false;
CRITERIAROW:
      for (j=1; j<criteriainfo.nrows; j++) { // go through each criteria row
         for (k=0; k<criteriainfo.ncols; k++) { // look at each column
            criteriacr = SocialCalc.crToCoord(criteriainfo.col1num + k, criteriainfo.row1num + j); // where criteria is
            cell = criteriainfo.sheetdata.GetAssuredCell(criteriacr);
            criteria = cell.datavalue;
            if (typeof criteria == "string" && criteria.length == 0) continue; // blank items are OK
            testcol = criteriafieldnums[k];
            testcr = SocialCalc.crToCoord(testcol, dbinfo.row1num + i); // cell to check
            cell = criteriainfo.sheetdata.GetAssuredCell(testcr);
            if (!scf.TestCriteria(cell.datavalue, cell.valuetype || "b", criteria)) {
               continue CRITERIAROW; // does not meet criteria - check next row
               }
            }
         testok = true; // met all the criteria
         break CRITERIAROW;
         }
      if (!testok) {
         continue;
         }

      cr = SocialCalc.crToCoord(targetcol, dbinfo.row1num + i); // get cell of this row to do the function on
      cell = dbinfo.sheetdata.GetAssuredCell(cr);

      value1.value = cell.datavalue;
      value1.type = cell.valuetype;
      t = value1.type.charAt(0);
      if (t == "n") count += 1;
      if (t != "b") counta += 1;
      if (t == "b") countblank += 1;

      if (t == "n") {
         v1 = value1.value-0; // get it as a number
         sum += v1;
         product *= v1;
         maxval = (maxval!=undefined) ? (v1 > maxval ? v1 : maxval) : v1;
         minval = (minval!=undefined) ? (v1 < minval ? v1 : minval) : v1;
         if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
            mk1 = v1;
            sk1 = 0;
            }
         else { // Accumulate S sub 1 through n as per Knuth noted above
            mk = mk1 + (v1 - mk1) / count;
            sk = sk1 + (v1 - mk1) * (v1 - mk);
            sk1 = sk;
            mk1 = mk;
            }
         resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
         }
      else if (t == "e" && resulttypesum.charAt(0) != "e") {
         resulttypesum = value1.type;
         }
      }

   resulttypesum = resulttypesum || "n";

   switch (fname) {
      case "DSUM":
         PushOperand(resulttypesum, sum);
         break;

      case "DPRODUCT": // may handle cases with text differently than some other spreadsheets
         PushOperand(resulttypesum, product);
         break;

      case "DMIN":
         PushOperand(resulttypesum, minval || 0);
         break;

      case "DMAX":
         PushOperand(resulttypesum, maxval || 0);
         break;

      case "DCOUNT":
         PushOperand("n", count);
         break;

      case "DCOUNTA":
         PushOperand("n", counta);
         break;

      case "DAVERAGE":
         if (count > 0) {
            PushOperand(resulttypesum, sum/count);
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "DSTDEV":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "DSTDEVP":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / count));
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "DVAR":
         if (count > 1) {
            PushOperand(resulttypesum, sk / (count - 1));
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "DVARP":
         if (count > 1) {
            PushOperand(resulttypesum, sk / count);
            }
         else {
            PushOperand("e#DIV/0!", 0);
            }
         break;

      case "DGET":
         if (count == 1) {
            PushOperand(resulttypesum, sum);
            }
         else if (count == 0) {
            PushOperand("e#VALUE!", 0);
            }
         else {
            PushOperand("e#NUM!", 0);
            }
         break;

      }

   return;

   }

SocialCalc.Formula.FunctionList["DAVERAGE"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DCOUNT"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DCOUNTA"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DGET"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DMAX"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DMIN"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DPRODUCT"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DSTDEV"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DSTDEVP"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DSUM"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DVAR"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DVARP"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];

/*
#
# colnum = SocialCalc.Formula.FieldToColnum(sheet, col1num, ncols, row1num, fieldname, fieldtype)
#
# If fieldname is a number, uses it, otherwise looks up string in cells in row to find field number
#
# If not found, returns 0.
#
*/

SocialCalc.Formula.FieldToColnum = function(sheet, col1num, ncols, row1num, fieldname, fieldtype) {

   var colnum, cell, value;

   if (fieldtype.charAt(0) == "n") { // number - return it if legal
      colnum = fieldname - 0; // make sure a number
      if (colnum <= 0 || colnum > ncols) {
         return 0;
         }
      return Math.floor(colnum);
      }

   if (fieldtype.charAt(0) != "t") { // must be text otherwise
      return 0;
      }

   fieldname = fieldname ? fieldname.toLowerCase() : "";

   for (colnum=0; colnum < ncols; colnum++) { // look through column headers for a match
      cell = sheet.GetAssuredCell(SocialCalc.crToCoord(col1num+colnum, row1num));
      value = cell.datavalue;
      value = (value+"").toLowerCase(); // ignore case
      if (value == fieldname) { // match
         return colnum+1;
         }         
      }
   return 0; // looked at all and no match

   }


/*
#
# HLOOKUP(value, range, row, [rangelookup])
# VLOOKUP(value, range, col, [rangelookup])
# MATCH(value, range, [rangelookup])
#
*/

SocialCalc.Formula.LookupFunctions = function(fname, operand, foperand, sheet) {

   var lookupvalue, range, offset, rangelookup, offsetvalue, rangeinfo;
   var c, r, cincr, rincr, previousOK, csave, rsave, cell, value, valuetype, cr, lookupvalue;

   var scf = SocialCalc.Formula;
   var operand_value_and_type = scf.OperandValueAndType;
   var lookup_result_type = scf.LookupResultType;
   var typelookupplus = scf.TypeLookupTable.plus;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   lookupvalue = operand_value_and_type(sheet, foperand);
   if (typeof lookupvalue.value == "string") {
      lookupvalue.value = lookupvalue.value.toLowerCase();
      }

   range = scf.TopOfStackValueAndType(sheet, foperand);

   rangelookup = 1; // default to true or 1
   if (fname == "MATCH") {
      if (foperand.length) {
         rangelookup = scf.OperandAsNumber(sheet, foperand);
         if (rangelookup.type.charAt(0) != "n") {
            PushOperand("e#VALUE!", 0);
            return;
            }
         if (foperand.length) {
            scf.FunctionArgsError(fname, operand);
            return 0;
            }
         rangelookup = rangelookup.value - 0;
         }
      }
   else {
      offsetvalue = scf.OperandAsNumber(sheet, foperand);
      if (offsetvalue.type.charAt(0) != "n") {
         PushOperand("e#VALUE!", 0);
         return;
         }
      offsetvalue = Math.floor(offsetvalue.value);
      if (foperand.length) {
         rangelookup = scf.OperandAsNumber(sheet, foperand);
         if (rangelookup.type.charAt(0) != "n") {
            PushOperand("e#VALUE!", 0);
            return;
            }
         if (foperand.length) {
            scf.FunctionArgsError(fname, operand);
            return 0;
            }
         rangelookup = rangelookup.value ? 1 : 0; // convert to 1 or 0
         }
      }
   lookupvalue.type = lookupvalue.type.charAt(0); // only deal with general type
   if (lookupvalue.type == "n") { // if number, make sure a number
      lookupvalue.value = lookupvalue.value - 0;
      }

   if (range.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }

   rangeinfo = scf.DecodeRangeParts(sheet, range.value, range.type);
   if (!rangeinfo) {
      PushOperand("e#REF!", 0);
      return;
      }

   c = 0;
   r = 0;
   cincr = 0;
   rincr = 0;
   if (fname == "HLOOKUP") {
      cincr = 1;
      if (offsetvalue > rangeinfo.nrows) {
         PushOperand("e#REF!", 0);
         return;
         }
      }
   else if (fname == "VLOOKUP") {
      rincr = 1;
      if (offsetvalue > rangeinfo.ncols) {
         PushOperand("e#REF!", 0);
         return;
         }
      }
   else if (fname == "MATCH") {
      if (rangeinfo.ncols > 1) {
         if (rangeinfo.nrows > 1) {
            PushOperand("e#N/A", 0);
            return;
            }
         cincr = 1;
         }
      else {
         rincr = 1;
         }
      }
   else {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }
   if (offsetvalue < 1 && fname != "MATCH") {
      PushOperand("e#VALUE!", 0);
      return 0;
      }

   previousOK; // if 1, previous test was <. If 2, also this one wasn't

   while (1) {
      cr = SocialCalc.crToCoord(rangeinfo.col1num + c, rangeinfo.row1num + r);
      cell = rangeinfo.sheetdata.GetAssuredCell(cr);
      value = cell.datavalue;
      valuetype = cell.valuetype ? cell.valuetype.charAt(0) : "b"; // only deal with general types
      if (valuetype == "n") {
         value = value - 0; // make sure number
         }
      if (rangelookup) { // rangelookup type 1 or -1: look for within brackets for matches
         if (lookupvalue.type == "n" && valuetype == "n") {
            if (lookupvalue.value == value) { // match
               break;
               }
            if ((rangelookup > 0 && lookupvalue.value > value)
                || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
               previousOK = 1;
               csave = c; // remember col and row of last OK
               rsave = r;
               }
            else if (previousOK) { // last one was OK, this one isn't
               previousOK = 2;
               break;
               }
            }

         else if (lookupvalue.type == "t" && valuetype == "t") {
            value = typeof value == "string" ? value.toLowerCase() : "";
            if (lookupvalue.value == value) { // match
               break;
               }
            if ((rangelookup > 0 && lookupvalue.value > value)
                || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
               previousOK = 1;
               csave = c;
               rsave = r;
               }
            else if (previousOK) { // last one was OK, this one isn't
               previousOK = 2;
               break;
               }
            }
         }
      else { // exact value matches
         if (lookupvalue.type == "n" && valuetype == "n") {
            if (lookupvalue.value == value) { // match
               break;
               }
            }
         else if (lookupvalue.type == "t" && valuetype == "t") {
            value = typeof value == "string" ? value.toLowerCase() : "";
            if (lookupvalue.value == value) { // match
               break;
               }
            }
         }

      r += rincr;
      c += cincr;
      if (r >= rangeinfo.nrows || c >= rangeinfo.ncols) { // end of range to check, no exact match
         if (previousOK) { // at least one could have been OK
            previousOK = 2;
            break;
            }
         PushOperand("e#N/A", 0);
         return;
         }
      }

   if (previousOK == 2) { // back to last OK
      r = rsave;
      c = csave;
      }

   if (fname == "MATCH") {
      value = c + r + 1; // only one may be <> 0
      valuetype = "n";
      }
   else {
      cr = SocialCalc.crToCoord(rangeinfo.col1num+c+(fname == "VLOOKUP" ? offsetvalue-1 : 0), rangeinfo.row1num+r+(fname == "HLOOKUP" ? offsetvalue-1 : 0));
      cell = rangeinfo.sheetdata.GetAssuredCell(cr);
      value = cell.datavalue;
      valuetype = cell.valuetype;
      }
   PushOperand(valuetype, value);

   return;

   }

SocialCalc.Formula.FunctionList["HLOOKUP"] = [SocialCalc.Formula.LookupFunctions, -3, "hlookup", "", "lookup"];
SocialCalc.Formula.FunctionList["MATCH"] = [SocialCalc.Formula.LookupFunctions, -2, "match", "", "lookup"];
SocialCalc.Formula.FunctionList["VLOOKUP"] = [SocialCalc.Formula.LookupFunctions, -3, "vlookup", "", "lookup"];

/*
#
# INDEX(range, rownum, colnum)
#
*/

SocialCalc.Formula.IndexFunction = function(fname, operand, foperand, sheet) {

   var range, sheetname, indexinfo, rowindex, colindex, result, resulttype;

   var scf = SocialCalc.Formula;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   range = scf.TopOfStackValueAndType(sheet, foperand); // get range
   if (range.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }
   indexinfo = scf.DecodeRangeParts(sheet, range.value, range.type);
   if (indexinfo.sheetname) {
      sheetname = "!" + indexinfo.sheetname;
      }
   else {
      sheetname = "";
      }

   rowindex = {value:0};
   colindex = {value:0};

   if (foperand.length) { // look for row number
      rowindex = scf.OperandAsNumber(sheet, foperand);
      if (rowindex.type.charAt(0) != "n" || rowindex.value < 0) {
         PushOperand("e#VALUE!", 0);
         return;
         }
      if (foperand.length) { // look for col number
         colindex = scf.OperandAsNumber(sheet, foperand);
         if (colindex.type.charAt(0) != "n" || colindex.value < 0) {
            PushOperand("e#VALUE!", 0);
            return;
            }
         if (foperand.length) {
            scf.FunctionArgsError(fname, operand);
            return 0;
            }
         }
      else { // col number missing
         if (indexinfo.nrows == 1) { // if only one row, then rowindex is really colindex
            colindex.value = rowindex.value;
            rowindex.value = 0;
            }
         }
      }

   if (rowindex.value > indexinfo.nrows || colindex.value > indexinfo.ncols) {
      PushOperand("e#REF!", 0);
      return;
      }

   if (rowindex.value == 0) {
      if (colindex.value == 0) {
         if (indexinfo.nrows == 1 && indexinfo.ncols == 1) {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname;
            resulttype = "coord";
            }
         else {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname + "|" +
                     SocialCalc.crToCoord(indexinfo.col1num+indexinfo.ncols-1, indexinfo.row1num+indexinfo.nrows-1) + 
                     "|";
            resulttype = "range";
            }
         }
      else {
         if (indexinfo.nrows == 1) {
            result = SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num) + sheetname;
            resulttype = "coord";
            }
         else {
            result = SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num) + sheetname + "|" +
                     SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num+indexinfo.nrows-1) +
                     "|";
            resulttype = "range";
            }
         }
      }
   else {
      if (colindex.value == 0) {
         if (indexinfo.ncols == 1) {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num+rowindex.value-1) + sheetname;
            resulttype = "coord";
            }
         else {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num+rowindex.value-1) + sheetname + "|" +
                     SocialCalc.crToCoord(indexinfo.col1num+indexinfo.ncols-1, indexinfo.row1num+rowindex.value-1) +
                     "|";
            resulttype = "range";
            }
         }
      else {
         result = SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num+rowindex.value-1) + sheetname;
         resulttype = "coord";
         }
      }

   PushOperand(resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["INDEX"] = [SocialCalc.Formula.IndexFunction, -1, "index", "", "lookup"];

/*
#
# COUNTIF(c1:c2,"criteria")
# SUMIF(c1:c2,"criteria",[range2])
#
*/

SocialCalc.Formula.CountifSumifFunctions = function(fname, operand, foperand, sheet) {

   var range, criteria, sumrange, f2operand, result, resulttype, value1, value2;
   var sum = 0;
   var resulttypesum = "";
   var count = 0;

   var scf = SocialCalc.Formula;
   var operand_value_and_type = scf.OperandValueAndType;
   var lookup_result_type = scf.LookupResultType;
   var typelookupplus = scf.TypeLookupTable.plus;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   range = scf.TopOfStackValueAndType(sheet, foperand); // get range or coord
   criteria = scf.OperandAsText(sheet, foperand); // get criteria
   if (fname == "SUMIF") {
      if (foperand.length == 1) { // three arg form of SUMIF
         sumrange = scf.TopOfStackValueAndType(sheet, foperand);
         }
      else if (foperand.length == 0) { // two arg form
         sumrange = {value: range.value, type: range.type};
         }
      else {
         scf.FunctionArgsError(fname, operand);
         return 0;
         }
      }
   else {
      sumrange = {value: range.value, type: range.type};
      }

   if (criteria.type.charAt(0) == "n") {
      criteria.value = criteria.value + ""; // make text
      }
   else if (criteria.type.charAt(0) == "e") { // error
      criteria.value = null;
      }
   else if (criteria.type.charAt(0) == "b") { // blank here is undefined
      criteria.value = null;
      }

   if (range.type != "coord" && range.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }

   if (fname == "SUMIF" && sumrange.type != "coord" && sumrange.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }

   foperand.push(range);
   f2operand = []; // to allow for 3 arg form
   f2operand.push(sumrange);

   while (foperand.length) {
      value1 = operand_value_and_type(sheet, foperand);
      value2 = operand_value_and_type(sheet, f2operand);

      if (!scf.TestCriteria(value1.value, value1.type, criteria.value)) {
         continue;
         }

      count += 1;

      if (value2.type.charAt(0) == "n") {
         sum += value2.value-0;
         resulttypesum = lookup_result_type(value2.type, resulttypesum || value2.type, typelookupplus);
         }
      else if (value2.type.charAt(0) == "e" && resulttypesum.charAt(0) != "e") {
         resulttypesum = value2.type;
         }
      }

   resulttypesum = resulttypesum || "n";

   if (fname == "SUMIF") {
      PushOperand(resulttypesum, sum);
      }
   else if (fname == "COUNTIF") {
      PushOperand("n", count);
      }

   return;

   }

SocialCalc.Formula.FunctionList["COUNTIF"] = [SocialCalc.Formula.CountifSumifFunctions, 2, "rangec", "", "stat"];
SocialCalc.Formula.FunctionList["SUMIF"] = [SocialCalc.Formula.CountifSumifFunctions, -2, "sumif", "", "stat"];

/*
#
# SUMIFS(c1:c2, c3:c4,"criteria", [c5:c6,"criteria", ...])
#
*/

SocialCalc.Formula.SumifsFunction = function(fname, operand, foperand, sheet) {
   var range, criteria, sumrange, f2operand, result, resulttype, value1, value2;
   var sum = 0;
   var resulttypesum = "";
   var count = 0;

   var scf = SocialCalc.Formula;
   var operand_value_and_type = scf.OperandValueAndType;
   var lookup_result_type = scf.LookupResultType;
   var typelookupplus = scf.TypeLookupTable.plus;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   sumrange = scf.TopOfStackValueAndType(sheet, foperand);
   if (sumrange.type != "coord" && sumrange.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }

   var ranges = [], criterias = [];
   while (foperand.length) {
      range = scf.TopOfStackValueAndType(sheet, foperand); // get range or coord
      criteria = scf.OperandAsText(sheet, foperand); // get criteria
      if (criteria.type.charAt(0) == "n") {
         criteria.value = criteria.value + ""; // make text
         }
      else if (criteria.type.charAt(0) == "e") { // error
         criteria.value = null;
         }
      else if (criteria.type.charAt(0) == "b") { // blank here is undefined
         criteria.value = null;
         }
      if (range.type != "coord" && range.type != "range") {
         scf.FunctionArgsError(fname, operand);
         return 0;
         }
      ranges.push([range]);
      criterias.push(criteria);
      }

      f2operand = [];
      f2operand.push(sumrange);

   while (f2operand.length) {
      value2 = operand_value_and_type(sheet, f2operand);

      var all_good = true;
      for (var i=0; i < ranges.length; i++) {
         value1 = operand_value_and_type(sheet, ranges[i]);
         if (!scf.TestCriteria(value1.value, value1.type, criterias[i].value)) {
            all_good = false;
            break;
            }
         }
      if (!all_good) { continue; }

      if (value2.type.charAt(0) == "n") {
         sum += value2.value-0;
         resulttypesum = lookup_result_type(value2.type, resulttypesum || value2.type, typelookupplus);
         }
      else if (value2.type.charAt(0) == "e" && resulttypesum.charAt(0) != "e") {
         resulttypesum = value2.type;
         }
      }

   resulttypesum = resulttypesum || "n";
   PushOperand(resulttypesum, sum);
   return;

   }


SocialCalc.Formula.FunctionList["SUMIFS"] = [SocialCalc.Formula.SumifsFunction, -3, "sumifs", "", "stat"];

/*
#
# IF(cond,truevalue,falsevalue)
#
*/

SocialCalc.Formula.IfFunction = function(fname, operand, foperand, sheet) {

   var cond, t;

   cond = SocialCalc.Formula.OperandValueAndType(sheet, foperand);
   t = cond.type.charAt(0);
   if (t != "n" && t != "b") {
      operand.push({type: "e#VALUE!", value: 0});
      return;
      }

   var op1, op2;

   op1 = foperand.pop();
   if (foperand.length == 1) {
      op2 = foperand.pop();
      }
   else if (foperand.length == 0) {
      op2 = {type: "n", value: 0};
      }
   else {
      scf.FunctionArgsError(fname, operand);
      return;
   }

   operand.push(cond.value ? op1 : op2);

   }

// Add to function list
SocialCalc.Formula.FunctionList["IF"] = [SocialCalc.Formula.IfFunction, -2, "iffunc", "", "test"];

/*
#
# DATE(year,month,day)
#
*/

SocialCalc.Formula.DateFunction = function(fname, operand, foperand, sheet) {

   var scf = SocialCalc.Formula;
   var result = 0;
   var year = scf.OperandAsNumber(sheet, foperand);
   var month = scf.OperandAsNumber(sheet, foperand);
   var day = scf.OperandAsNumber(sheet, foperand);
   var resulttype = scf.LookupResultType(year.type, month.type, scf.TypeLookupTable.twoargnumeric);
   resulttype = scf.LookupResultType(resulttype, day.type, scf.TypeLookupTable.twoargnumeric);
   if (resulttype.charAt(0) == "n") {
      result = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(
                  Math.floor(year.value), Math.floor(month.value), Math.floor(day.value)
                  ) - SocialCalc.FormatNumber.datevalues.julian_offset;
      resulttype = "nd";
      }
   scf.PushOperand(operand, resulttype, result);
   return;

   }

SocialCalc.Formula.FunctionList["DATE"] = [SocialCalc.Formula.DateFunction, 3, "date", "", "datetime"];

/*
#
# TIME(hour,minute,second)
#
*/

SocialCalc.Formula.TimeFunction = function(fname, operand, foperand, sheet) {

   var scf = SocialCalc.Formula;
   var result = 0;
   var hours = scf.OperandAsNumber(sheet, foperand);
   var minutes = scf.OperandAsNumber(sheet, foperand);
   var seconds = scf.OperandAsNumber(sheet, foperand);
   var resulttype = scf.LookupResultType(hours.type, minutes.type, scf.TypeLookupTable.twoargnumeric);
   resulttype = scf.LookupResultType(resulttype, seconds.type, scf.TypeLookupTable.twoargnumeric);
   if (resulttype.charAt(0) == "n") {
      result = ((hours.value * 60 * 60) + (minutes.value * 60) + seconds.value) / (24*60*60);
      resulttype = "nt";
      }
   scf.PushOperand(operand, resulttype, result);
   return;

   }

SocialCalc.Formula.FunctionList["TIME"] = [SocialCalc.Formula.TimeFunction, 3, "hms", "", "datetime"];

/*
#
# DAY(date)
# MONTH(date)
# YEAR(date)
# WEEKDAY(date, [type])
#
*/

SocialCalc.Formula.DMYFunctions = function(fname, operand, foperand, sheet) {

   var ymd, dtype, doffset;
   var scf = SocialCalc.Formula;
   var result = 0;

   var datevalue = scf.OperandAsNumber(sheet, foperand);
   var resulttype = scf.LookupResultType(datevalue.type, datevalue.type, scf.TypeLookupTable.oneargnumeric);

   if (resulttype.charAt(0) == "n") {
      ymd = SocialCalc.FormatNumber.convert_date_julian_to_gregorian(
               Math.floor(datevalue.value + SocialCalc.FormatNumber.datevalues.julian_offset));
      switch (fname) {
         case "DAY":
            result = ymd.day;
            break;

         case "MONTH":
            result = ymd.month;
            break;

         case "YEAR":
            result = ymd.year;
            break;

         case "WEEKDAY":
            dtype = {value: 1};
            if (foperand.length) { // get type if present
               dtype = scf.OperandAsNumber(sheet, foperand);
               if (dtype.type.charAt(0) != "n" || dtype.value < 1 || dtype.value > 3) {
                  scf.PushOperand(operand, "e#VALUE!", 0);
                  return;
                  }
               if (foperand.length) { // extra args
                  scf.FunctionArgsError(fname, operand);
                  return;
                  }
               }
            doffset = 6;
            if (dtype.value > 1) {
               doffset -= 1;
               }
            result = Math.floor(datevalue.value+doffset) % 7 + (dtype.value < 3 ? 1 : 0);
            break;
         }
      }

   scf.PushOperand(operand, resulttype, result);
   return;

   }

SocialCalc.Formula.FunctionList["DAY"] = [SocialCalc.Formula.DMYFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["MONTH"] = [SocialCalc.Formula.DMYFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["YEAR"] = [SocialCalc.Formula.DMYFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["WEEKDAY"] = [SocialCalc.Formula.DMYFunctions, -1, "weekday", "", "datetime"];

/*
#
# HOUR(datetime)
# MINUTE(datetime)
# SECOND(datetime)
#
*/

SocialCalc.Formula.HMSFunctions = function(fname, operand, foperand, sheet) {

   var hours, minutes, seconds, fraction;
   var scf = SocialCalc.Formula;
   var result = 0;

   var datetime = scf.OperandAsNumber(sheet, foperand);
   var resulttype = scf.LookupResultType(datetime.type, datetime.type, scf.TypeLookupTable.oneargnumeric);

   if (resulttype.charAt(0) == "n") {
      if (datetime.value < 0) {
         scf.PushOperand(operand, "e#NUM!", 0); // must be non-negative
         return;
         }
      fraction = datetime.value - Math.floor(datetime.value); // fraction of a day
      fraction *= 24;
      hours = Math.floor(fraction);
      fraction -= Math.floor(fraction);
      fraction *= 60;
      minutes = Math.floor(fraction);
      fraction -= Math.floor(fraction);
      fraction *= 60;
      seconds = Math.floor(fraction + (datetime.value >= 0 ? 0.5: -0.5));
      if (fname == "HOUR") {
         result = hours;
         }
      else if (fname == "MINUTE") {
         result = minutes;
         }
      else if (fname == "SECOND") {
         result = seconds;
         }
      }

   scf.PushOperand(operand, resulttype, result);
   return;

   }

SocialCalc.Formula.FunctionList["HOUR"] = [SocialCalc.Formula.HMSFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["MINUTE"] = [SocialCalc.Formula.HMSFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["SECOND"] = [SocialCalc.Formula.HMSFunctions, 1, "v", "", "datetime"];

/*
#
# EXACT(v1,v2)
#
*/

SocialCalc.Formula.ExactFunction = function(fname, operand, foperand, sheet) {

   var scf = SocialCalc.Formula;
   var result = 0;
   var resulttype = "nl";

   var value1 = scf.OperandValueAndType(sheet, foperand);
   var v1type = value1.type.charAt(0);
   var value2 = scf.OperandValueAndType(sheet, foperand);
   var v2type = value2.type.charAt(0);

   if (v1type == "t") {
      if (v2type == "t") {
         result = value1.value == value2.value ? 1 : 0;
         }
      else if (v2type == "b") {
         result = value1.value.length ? 0 : 1;
         }
      else if (v2type == "n") {
         result = value1.value == value2.value+"" ? 1 : 0;
         }
      else if (v2type == "e") {
         result = value2.value;
         resulttype = value2.type;
         }
      else {
         result = 0;
         }
      }
   else if (v1type == "n") {
      if (v2type == "n") {
         result = value1.value-0 == value2.value-0 ? 1 : 0;
         }
      else if (v2type == "b") {
         result = 0;
         }
      else if (v2type == "t") {
         result = value1.value+"" == value2.value ? 1 : 0;
         }
      else if (v2type == "e") {
         result = value2.value;
         resulttype = value2.type;
         }
      else {
         result = 0;
         }
      }
   else if (v1type == "b") {
      if (v2type == "t") {
         result = value2.value.length ? 0 : 1;
         }
      else if (v2type == "b") {
         result = 1;
         }
      else if (v2type == "n") {
         result = 0;
         }
      else if (v2type == "e") {
         result = value2.value;
         resulttype = value2.type;
         }
      else {
         result = 0;
         }
      }
   else if (v1type == "e") {
      result = value1.value;
      resulttype = value1.type;
      }

   scf.PushOperand(operand, resulttype, result);
   return;

   }

SocialCalc.Formula.FunctionList["EXACT"] = [SocialCalc.Formula.ExactFunction, 2, "", "", "text"];

/*
#
# FIND(key,string,[start])
# LEFT(string,[length])
# LEN(string)
# LOWER(string)
# MID(string,start,length)
# PROPER(string)
# REPLACE(string,start,length,new)
# REPT(string,count)
# RIGHT(string,[length])
# SUBSTITUTE(string,old,new,[which])
# TRIM(string)
# HEXCODE(string)
# UPPER(string)
#
*/

// SocialCalc.Formula.ArgList has an array for each function, one entry for each possible arg (up to max).
// Min args are specified in SocialCalc.Formula.FunctionList.
// If array element is 1 then it's a text argument, if it's 0 then it's numeric, if -1 then just get whatever's there
// Text values are manipulated as UTF-8, converting from and back to byte strings

SocialCalc.Formula.ArgList = {
                FIND: [1, 1, 0],
                LEFT: [1, 0],
                LEN: [1],
                LOWER: [1],
                MID: [1, 0, 0],
                PROPER: [1],
                REPLACE: [1, 0, 0, 1],
                REPT: [1, 0],
                RIGHT: [1, 0],
                SUBSTITUTE: [1, 1, 1, 0],
                TRIM: [1],
                HEXCODE: [1],
                UPPER: [1]
               };

SocialCalc.Formula.StringFunctions = function(fname, operand, foperand, sheet) {

   var i, value, offset, len, start, count;
   var scf = SocialCalc.Formula;
   var result = 0;
   var resulttype = "e#VALUE!";

   var numargs = foperand.length;
   var argdef = scf.ArgList[fname];
   var operand_value = [];
   var operand_type = [];

   for (i=1; i <= numargs; i++) { // go through each arg, get value and type, and check for errors
      if (i > argdef.length) { // too many args
         scf.FunctionArgsError(fname, operand);
         return;
         }
      if (argdef[i-1] == 0) {
         value = scf.OperandAsNumber(sheet, foperand);
         }
      else if (argdef[i-1] == 1) {
         value = scf.OperandAsText(sheet, foperand);
         }
      else if (argdef[i-1] == -1) {
         value = scf.OperandValueAndType(sheet, foperand);
         }
      operand_value[i] = value.value;
      operand_type[i] = value.type;
      if (value.type.charAt(0) == "e") {
         scf.PushOperand(operand, value.type, result);
         return;
         }
      }

   switch (fname) {
      case "FIND":
         offset = operand_type[3] ? operand_value[3]-1 : 0;
         if (offset < 0) {
            result = "Start is before string"; // !! not displayed, no need to translate
            }
         else {
            result = operand_value[2].indexOf(operand_value[1], offset); // (null string matches first char)
            if (result >= 0) {
               result += 1;
               resulttype = "n";
               }
            else {
               result = "Not found"; // !! not displayed, error is e#VALUE!
               }
            }
         break;

      case "LEFT":
         len = operand_type[2] ? operand_value[2]-0 : 1;
         if (len < 0) {
            result = "Negative length";
            }
         else {
            result = operand_value[1].substring(0, len);
            resulttype = "t";
            }
         break;

      case "LEN":
         result = operand_value[1].length;
         resulttype = "n";
         break;

      case "LOWER":
         result = operand_value[1].toLowerCase();
         resulttype = "t";
         break;

      case "MID":
         start = operand_value[2]-0;
         len = operand_value[3]-0;
         if (len < 1 || start < 1) {
            result = "Bad arguments";
            }
         else {
            result = operand_value[1].substring(start-1, start+len-1);
            resulttype = "t";
            }
         break;

      case "PROPER":
         result = operand_value[1].replace(/\b\w+\b/g, function(word) {
                     return word.substring(0,1).toUpperCase() + 
                        word.substring(1);
                     }); // uppercase first character of words (see JavaScript, Flanagan, 5th edition, page 704)
         resulttype = "t";
         break;

      case "REPLACE":
         start = operand_value[2]-0;
         len = operand_value[3]-0;
         if (len < 0 || start < 1) {
            result = "Bad arguments";
            }
         else {
            result = operand_value[1].substring(0, start-1) + operand_value[4] + 
               operand_value[1].substring(start-1+len);
            resulttype = "t";
            }
         break;

      case "REPT":
         count = operand_value[2]-0;
         if (count < 0) {
            result = "Negative count";
            }
         else {
            result = "";
            for (; count > 0; count--) {
               result += operand_value[1];
               }
            resulttype = "t";
            }
         break;

      case "RIGHT":
         len = operand_type[2] ? operand_value[2]-0 : 1;
         if (len < 0) {
            result = "Negative length";
            }
         else {
            result = operand_value[1].slice(-len);
            resulttype = "t";
            }
         break;

      case "SUBSTITUTE":
         fulltext = operand_value[1];
         oldtext = operand_value[2];
         newtext = operand_value[3];
         if (operand_value[4] != null) {
            which = operand_value[4]-0;
            if (which <= 0) {
               result = "Non-positive instance number";
               break;
               }
            }
         else {
            which = 0;
            }
         count = 0;
         oldpos = 0;
         result = "";
         while (true) {
            pos = fulltext.indexOf(oldtext, oldpos);
            if (pos >= 0) {
               count++; //!!!!!! old test just in case: if (count>1000) {alert(pos); break;}
               result += fulltext.substring(oldpos, pos);
               if (which==0) {
                  result += newtext; // substitute
                  }
               else if (which==count) {
                  result += newtext + fulltext.substring(pos+oldtext.length);
                  break;
                  }
               else {
                  result += oldtext; // leave as was
                  }
               oldpos = pos + oldtext.length;
               }
            else { // no more
               result += fulltext.substring(oldpos);
               break;
               }
            }
         resulttype = "t";
         break;

      case "TRIM":
         result = operand_value[1];
         result = result.replace(/^ */, "");
         result = result.replace(/ *$/, "");
         result = result.replace(/ +/g, " ");
         resulttype = "t";
         break;

      case "HEXCODE":
         result = String(operand_value[1]);
         var code = result.charCodeAt(0);
         if (0xD800 <= code && code <= 0xDBFF) {
             var next = result.charCodeAt(1);
             if (0xDC00 <= next && next <= 0xDFFF) {
                 code = ((code - 0xD800) * 0x400) + (next - 0xDC00) + 0x10000;
             }
         }
         result = code.toString(16).toUpperCase();
         resulttype = "t";
         break;

      case "UPPER":
         result = operand_value[1].toUpperCase();
         resulttype = "t";
         break;

      }

   scf.PushOperand(operand, resulttype, result);
   return;

   }

SocialCalc.Formula.FunctionList["FIND"] = [SocialCalc.Formula.StringFunctions, -2, "find", "", "text"];
SocialCalc.Formula.FunctionList["LEFT"] = [SocialCalc.Formula.StringFunctions, -2, "tc", "", "text"];
SocialCalc.Formula.FunctionList["LEN"] = [SocialCalc.Formula.StringFunctions, 1, "txt", "", "text"];
SocialCalc.Formula.FunctionList["LOWER"] = [SocialCalc.Formula.StringFunctions, 1, "txt", "", "text"];
SocialCalc.Formula.FunctionList["MID"] = [SocialCalc.Formula.StringFunctions, 3, "mid", "", "text"];
SocialCalc.Formula.FunctionList["PROPER"] = [SocialCalc.Formula.StringFunctions, 1, "v", "", "text"];
SocialCalc.Formula.FunctionList["REPLACE"] = [SocialCalc.Formula.StringFunctions, 4, "replace", "", "text"];
SocialCalc.Formula.FunctionList["REPT"] = [SocialCalc.Formula.StringFunctions, 2, "tc", "", "text"];
SocialCalc.Formula.FunctionList["RIGHT"] = [SocialCalc.Formula.StringFunctions, -1, "tc", "", "text"];
SocialCalc.Formula.FunctionList["SUBSTITUTE"] = [SocialCalc.Formula.StringFunctions, -3, "subs", "", "text"];
SocialCalc.Formula.FunctionList["TRIM"] = [SocialCalc.Formula.StringFunctions, 1, "v", "", "text"];
SocialCalc.Formula.FunctionList["HEXCODE"] = [SocialCalc.Formula.StringFunctions, 1, "v", "", "text"];
SocialCalc.Formula.FunctionList["UPPER"] = [SocialCalc.Formula.StringFunctions, 1, "v", "", "text"];

/*
#
# is_functions:
#
# ISBLANK(value)
# ISERR(value)
# ISERROR(value)
# ISLOGICAL(value)
# ISNA(value)
# ISNONTEXT(value)
# ISNUMBER(value)
# ISTEXT(value)
#
*/

SocialCalc.Formula.IsFunctions = function(fname, operand, foperand, sheet) {

   var scf = SocialCalc.Formula;
   var result = 0;
   var resulttype = "nl";

   var value = scf.OperandValueAndType(sheet, foperand);
   var t = value.type.charAt(0);

   switch (fname) {

      case "ISBLANK":
         result = value.type == "b" ? 1 : 0;
         break;

      case "ISERR":
         result = t == "e" ? (value.type == "e#N/A" ? 0 : 1) : 0;
         break;

      case "ISERROR":
         result = t == "e" ? 1 : 0;
         break;

      case "ISLOGICAL":
         result = value.type == "nl" ? 1 : 0;
         break;

      case "ISNA":
         result = value.type == "e#N/A" ? 1 : 0;
         break;

      case "ISNONTEXT":
         result = t == "t" ? 0 : 1;
         break;

      case "ISNUMBER":
         result = t == "n" ? 1 : 0;
         break;

      case "ISTEXT":
         result = t == "t" ? 1 : 0;
         break;
      }

   scf.PushOperand(operand, resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["ISBLANK"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISERR"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISERROR"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISLOGICAL"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISNA"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISNONTEXT"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISNUMBER"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISTEXT"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];

/*
#
# ntv_functions:
#
# N(value)
# T(value)
# VALUE(value)
#
*/

SocialCalc.Formula.NTVFunctions = function(fname, operand, foperand, sheet) {

   var scf = SocialCalc.Formula;
   var result = 0;
   var resulttype = "e#VALUE!";

   var value = scf.OperandValueAndType(sheet, foperand);
   var t = value.type.charAt(0);

   switch (fname) {

      case "N":
         result = t == "n" ? value.value-0 : 0;
         resulttype = "n";
         break;

      case "T":
         result = t == "t" ? value.value+"" : "";
         resulttype = "t";
         break;

      case "VALUE":
         if (t == "n" || t == "b") {
            result = value.value || 0;
            resulttype = "n";
            }
         else if (t == "t") {
            value = SocialCalc.DetermineValueType(value.value);
            if (value.type.charAt(0) != "n") {
               result = 0;
               resulttype = "e#VALUE!";
               }
            else {
               result = value.value-0;
               resulttype = "n";
               }
            }
         break;
      }

   if (t == "e") { // error trumps
      resulttype = value.type;
      }

   scf.PushOperand(operand, resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["N"] = [SocialCalc.Formula.NTVFunctions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["T"] = [SocialCalc.Formula.NTVFunctions, 1, "v", "", "text"];
SocialCalc.Formula.FunctionList["VALUE"] = [SocialCalc.Formula.NTVFunctions, 1, "v", "", "text"];

/*
#
# ABS(value)
# ACOS(value)
# ASIN(value)
# ATAN(value)
# COS(value)
# DEGREES(value)
# EVEN(value)
# EXP(value)
# FACT(value)
# INT(value)
# LN(value)
# LOG10(value)
# ODD(value)
# RADIANS(value)
# SIN(value)
# SQRT(value)
# TAN(value)
#
*/

SocialCalc.Formula.Math1Functions = function(fname, operand, foperand, sheet) {

   var v1, value, f;
   var result = {};

   var scf = SocialCalc.Formula;

   v1 = scf.OperandAsNumber(sheet, foperand);
   value = v1.value;
   result.type = scf.LookupResultType(v1.type, v1.type, scf.TypeLookupTable.oneargnumeric);

   if (result.type == "n") {
      switch (fname) {
         case "ABS":
            value = Math.abs(value);
            break;

         case "ACOS":
            if (value >= -1 && value <= 1) {
               value = Math.acos(value);
               }
            else {
               result.type = "e#NUM!";
               }
            break;

         case "ASIN":
            if (value >= -1 && value <= 1) {
               value = Math.asin(value);
               }
            else {
               result.type = "e#NUM!";
               }
            break;

         case "ATAN":
            value = Math.atan(value);
            break;

         case "COS":
            value = Math.cos(value);
            break;

         case "DEGREES":
            value = value * 180/Math.PI;
            break;

         case "EVEN":
            value = value < 0 ? -value : value;
            if (value != Math.floor(value)) {
               value = Math.floor(value + 1) + (Math.floor(value + 1) % 2);
               }
            else { // integer
               value = value + (value % 2);
               }
            if (v1.value < 0) value = -value;
            break;

         case "EXP":
            value = Math.exp(value);
            break;

         case "FACT":
            f = 1;
            value = Math.floor(value);
            for (;value>0;value--) {
               f *= value;
               }
            value = f;
            break;

         case "INT":
            value = Math.floor(value); // spreadsheet INT is floor(), not int()
            break;

         case "LN":
            if (value <= 0) {
               result.type = "e#NUM!";
               result.error = SocialCalc.Constants.s_sheetfunclnarg;
               }
            value = Math.log(value);
            break;

         case "LOG10":
            if (value <= 0) {
               result.type = "e#NUM!";
               result.error = SocialCalc.Constants.s_sheetfunclog10arg;
               }
            value = Math.log(value)/Math.log(10);
            break;

         case "ODD":
            value = value < 0 ? -value : value;
            if (value != Math.floor(value)) {
               value = Math.floor(value + 1) + (1 - (Math.floor(value + 1) % 2));
               }
            else { // integer
               value = value + (1 - (value % 2));
               }
            if (v1.value < 0) value = -value;
            break;

         case "RADIANS":
            value = value * Math.PI/180;
            break;

         case "SIN":
            value = Math.sin(value);
            break;

         case "SQRT":
            if (value >= 0) {
               value = Math.sqrt(value);
               }
            else {
               result.type = "e#NUM!";
               }
            break;

         case "TAN":
            if (Math.cos(value) != 0) {
               value = Math.tan(value);
               }
            else {
               result.type = "e#NUM!";
               }
            break;
         }
      }

   result.value = value;
   operand.push(result);

   return null;

   }

// Add to function list
SocialCalc.Formula.FunctionList["ABS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ACOS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ASIN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ATAN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["COS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["DEGREES"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["EVEN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["EXP"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["FACT"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["INT"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["LN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["LOG10"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ODD"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["RADIANS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["SIN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["SQRT"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["TAN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];


/*
#
# ATAN2(x, y)
# MOD(a, b)
# POWER(a, b)
# TRUNC(value, precision)
#
*/

SocialCalc.Formula.Math2Functions = function(fname, operand, foperand, sheet) {

   var xval, yval, value, quotient, decimalscale, i;
   var result = {};

   var scf = SocialCalc.Formula;

   xval = scf.OperandAsNumber(sheet, foperand);
   yval = scf.OperandAsNumber(sheet, foperand);
   value = 0;
   result.type = scf.LookupResultType(xval.type, yval.type, scf.TypeLookupTable.twoargnumeric);

   if (result.type == "n") {
      switch (fname) {
         case "ATAN2":
            if (xval.value == 0 && yval.value == 0) {
               result.type = "e#DIV/0!";
               }
            else {
               result.value = Math.atan2(yval.value, xval.value);
               }
            break;

         case "POWER":
            result.value = Math.pow(xval.value, yval.value);
            if (isNaN(result.value)) {
               result.value = 0;
               result.type = "e#NUM!";
               }
            break;

         case "MOD": // en.wikipedia.org/wiki/Modulo_operation, etc.
            if (yval.value == 0) {
               result.type = "e#DIV/0!";
               }
            else {
               quotient = xval.value/yval.value;
               quotient = Math.floor(quotient);
               result.value = xval.value - (quotient * yval.value);
               }
            break;

         case "TRUNC":
            decimalscale = 1; // cut down to required number of decimal digits
            if (yval.value >= 0) {
               yval.value = Math.floor(yval.value);
               for (i=0; i<yval.value; i++) {
                  decimalscale *= 10;
                  }
               result.value = Math.floor(Math.abs(xval.value) * decimalscale) / decimalscale;
               }
            else if (yval.value < 0) {
               yval.value = Math.floor(-yval.value);
               for (i=0; i<yval.value; i++) {
                  decimalscale *= 10;
                  }
               result.value = Math.floor(Math.abs(xval.value) / decimalscale) * decimalscale;
               }
            if (xval.value < 0) {
               result.value = -result.value;
               }
            }
         }
 
   operand.push(result);

   return null;

   }

// Add to function list
SocialCalc.Formula.FunctionList["ATAN2"] = [SocialCalc.Formula.Math2Functions, 2, "xy", "", "math"];
SocialCalc.Formula.FunctionList["MOD"] = [SocialCalc.Formula.Math2Functions, 2, "", "", "math"];
SocialCalc.Formula.FunctionList["POWER"] = [SocialCalc.Formula.Math2Functions, 2, "", "", "math"];
SocialCalc.Formula.FunctionList["TRUNC"] = [SocialCalc.Formula.Math2Functions, 2, "valpre", "", "math"];

/*
#
# LOG(value,[base])
#
*/

SocialCalc.Formula.LogFunction = function(fname, operand, foperand, sheet) {

   var value, value2;
   var result = {};

   var scf = SocialCalc.Formula;

   result.value = 0;

   value = scf.OperandAsNumber(sheet, foperand);
   result.type = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);
   if (foperand.length == 1) {
      value2 = scf.OperandAsNumber(sheet, foperand);
      if (value2.type.charAt(0) != "n" || value2.value <= 0) {
         scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogsecondarg);
         return 0;
         }
      }
   else if (foperand.length != 0) {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }
   else {
      value2 = {value: Math.E, type: "n"};
      }

   if (result.type == "n") {
      if (value.value <= 0) {
         scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogfirstarg);
         return 0;
         }
      result.value = Math.log(value.value)/Math.log(value2.value);
      }

   operand.push(result);

   return;

   }

SocialCalc.Formula.FunctionList["LOG"] = [SocialCalc.Formula.LogFunction, -1, "log", "", "math"];


/*
#
# ROUND(value,[precision])
#
*/

SocialCalc.Formula.RoundFunction = function(fname, operand, foperand, sheet) {

   var value2, decimalscale, scaledvalue, i;

   var scf = SocialCalc.Formula;
   var result = 0;
   var resulttype = "e#VALUE!";

   var value = scf.OperandValueAndType(sheet, foperand);
   var resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);

   if (foperand.length == 1) {
      value2 = scf.OperandValueAndType(sheet, foperand);
      if (value2.type.charAt(0) != "n") {
         scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncroundsecondarg);
         return 0;
         }
      }
   else if (foperand.length != 0) {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }
   else {
      value2 = {value: 0, type: "n"}; // if no second arg, assume 0 for simple round
      }

   if (resulttype == "n") {
      value2.value = value2.value-0;
      if (value2.value == 0) {
         result = Math.round(value.value);
         }
      else if (value2.value > 0) {
         decimalscale = 1; // cut down to required number of decimal digits
         value2.value = Math.floor(value2.value);
         for (i=0; i<value2.value; i++) {
            decimalscale *= 10;
            }
         scaledvalue = Math.round(value.value * decimalscale);
         result = scaledvalue / decimalscale;
         }
      else if (value2.value < 0) {
         decimalscale = 1; // cut down to required number of decimal digits
         value2.value = Math.floor(-value2.value);
         for (i=0; i<value2.value; i++) {
            decimalscale *= 10;
            }
         scaledvalue = Math.round(value.value / decimalscale);
         result = scaledvalue * decimalscale;
         }
      }

   scf.PushOperand(operand, resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["ROUND"] = [SocialCalc.Formula.RoundFunction, -1, "vp", "", "math"];

/*
#
# CEILING(value, [significance])
# FLOOR(value, [significance])
#
*/

SocialCalc.Formula.CeilingFloorFunctions = function(fname, operand, foperand, sheet) {

   var scf = SocialCalc.Formula;
   var val, sig, t;

   var PushOperand = function(t, v) {operand.push({type: t, value: v});};

   val = scf.OperandValueAndType(sheet, foperand);
   t = val.type.charAt(0);
   if (t != "n") {
      PushOperand("e#VALUE!", 0);
      return;
      }
   if (val.value == 0) {
      PushOperand("n", 0);
      return;
      }

   if (foperand.length == 1) {
      sig = scf.OperandValueAndType(sheet, foperand);
      t = val.type.charAt(0);
      if (t != "n") {
         PushOperand("e#VALUE!", 0);
         return;
         }
      }
   else if (foperand.length == 0) {
      sig = {type: "n", value: val.value > 0 ? 1 : -1};
      }
   else {
      PushOperand("e#VALUE!", 0);
      return;
      }
   if (sig.value == 0) {
      PushOperand("n", 0);
      return;
      }
   if (sig.value * val.value < 0) {
      PushOperand("e#NUM!", 0);
      return;
      }

   switch (fname) {
      case "CEILING":
         PushOperand("n", Math.ceil(val.value / sig.value) * sig.value);
         break;
      case "FLOOR":
         PushOperand("n", Math.floor(val.value / sig.value) * sig.value);
         break;
      }

   return;

   }

SocialCalc.Formula.FunctionList["CEILING"] = [SocialCalc.Formula.CeilingFloorFunctions, -1, "vsig", "", "math"];
SocialCalc.Formula.FunctionList["FLOOR"] = [SocialCalc.Formula.CeilingFloorFunctions, -1, "vsig", "", "math"];

/*
#
# AND(v1,c1:c2,...)
# OR(v1,c1:c2,...)
#
*/

SocialCalc.Formula.AndOrFunctions = function(fname, operand, foperand, sheet) {

   var value1, result;

   var scf = SocialCalc.Formula;
   var resulttype = "";

   if (fname == "AND") {
      result = 1;
      }
   else if (fname == "OR") {
      result = 0;
      }

   while (foperand.length) {
      value1 = scf.OperandValueAndType(sheet, foperand);
      if (value1.type.charAt(0) == "n") {
         value1.value = value1.value-0;
         if (fname == "AND") {
            result = value1.value != 0 ? result : 0;
            }
         else if (fname == "OR") {
            result = value1.value != 0 ? 1 : result;
            }
         resulttype = scf.LookupResultType(value1.type, resulttype || "nl", scf.TypeLookupTable.propagateerror);
         }
      else if (value1.type.charAt(0) == "e" && resulttype.charAt(0) != "e") {
         resulttype = value1.type;
         }
      }
   if (resulttype.length < 1) {
      resulttype = "e#VALUE!";
      result = 0;
      }

   scf.PushOperand(operand, resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["AND"] = [SocialCalc.Formula.AndOrFunctions, -1, "vn", "", "test"];
SocialCalc.Formula.FunctionList["OR"] = [SocialCalc.Formula.AndOrFunctions, -1, "vn", "", "test"];

/*
#
# NOT(value)
#
*/

SocialCalc.Formula.NotFunction = function(fname, operand, foperand, sheet) {

   var result = 0;
   var scf = SocialCalc.Formula;
   var value = scf.OperandValueAndType(sheet, foperand);
   var resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.propagateerror);

   if (value.type.charAt(0) == "n" || value.type == "b") {
      result = value.value-0 != 0 ? 0 : 1; // do the "not" operation
      resulttype = "nl";
      }
   else if (value.type.charAt(0) == "t") {
      resulttype = "e#VALUE!";
      }

   scf.PushOperand(operand, resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["NOT"] = [SocialCalc.Formula.NotFunction, 1, "v", "", "test"];

/*
#
# CHOOSE(index,value1,value2,...)
#
*/

SocialCalc.Formula.ChooseFunction = function(fname, operand, foperand, sheet) {

   var resulttype, count, value1;
   var result = 0;
   var scf = SocialCalc.Formula;

   var cindex = scf.OperandAsNumber(sheet, foperand);

   if (cindex.type.charAt(0) != "n") {
      cindex.value = 0;
      }
   cindex.value = Math.floor(cindex.value);

   count = 0;
   while (foperand.length) {
      value1 = scf.TopOfStackValueAndType(sheet, foperand);
      count += 1;
      if (cindex.value == count) {
         result = value1.value;
         resulttype = value1.type;
         break;
         }
      }
   if (resulttype) { // found something
      scf.PushOperand(operand, resulttype, result);
      }
   else {
      scf.PushOperand(operand, "e#VALUE!", 0);
      }

   return;

   }

SocialCalc.Formula.FunctionList["CHOOSE"] = [SocialCalc.Formula.ChooseFunction, -2, "choose", "", "lookup"];

/*
#
# COLUMNS(c1:c2)
# ROWS(c1:c2)
#
*/

SocialCalc.Formula.ColumnsRowsFunctions = function(fname, operand, foperand, sheet) {

   var resulttype, rangeinfo;
   var result = 0;
   var scf = SocialCalc.Formula;

   var value1 = scf.TopOfStackValueAndType(sheet, foperand);

   if (value1.type == "coord") {
      result = 1;
      resulttype = "n";
      }

   else if (value1.type == "range") {
      rangeinfo = scf.DecodeRangeParts(sheet, value1.value);
      if (fname == "COLUMNS") {
         result = rangeinfo.ncols;
         }
      else if (fname == "ROWS") {
         result = rangeinfo.nrows;
         }
      resulttype = "n";
      }
   else {
      result = 0;
      resulttype = "e#VALUE!";
      }

   scf.PushOperand(operand, resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["COLUMNS"] = [SocialCalc.Formula.ColumnsRowsFunctions, 1, "range", "", "lookup"];
SocialCalc.Formula.FunctionList["ROWS"] = [SocialCalc.Formula.ColumnsRowsFunctions, 1, "range", "", "lookup"];


/*
#
# FALSE()
# NA()
# NOW()
# PI()
# TODAY()
# TRUE()
# RAND()
#
*/

SocialCalc.Formula.ZeroArgFunctions = function(fname, operand, foperand, sheet) {

   var startval, tzoffset, start_1_1_1970, seconds_in_a_day, nowdays;
   var result = {value: 0};

   switch (fname) {
      case "FALSE":
         result.type = "nl";
         result.value = 0;
         break;

      case "NA":
         result.type = "e#N/A";
         break;

      case "NOW":
         startval = new Date();
         tzoffset = startval.getTimezoneOffset();
         startval = startval.getTime() / 1000; // convert to seconds
         start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
         seconds_in_a_day = 24 * 60 * 60;
         nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset/(24*60);
         result.value = nowdays;
         result.type = "ndt";
         SocialCalc.Formula.FreshnessInfo.volatile.NOW = true; // remember
         break;

      case "PI":
         result.type = "n";
         result.value = Math.PI;
         break;

      case "TODAY":
         startval = new Date();
         tzoffset = startval.getTimezoneOffset();
         startval = startval.getTime() / 1000; // convert to seconds
         start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
         seconds_in_a_day = 24 * 60 * 60;
         nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset/(24*60);
         result.value = Math.floor(nowdays);
         result.type = "nd";
         SocialCalc.Formula.FreshnessInfo.volatile.TODAY = true; // remember
         break;

      case "TRUE":
         result.type = "nl";
         result.value = 1;
         break;

      case "RAND":
         result.type = "n";
         result.value = Math.random();
         SocialCalc.Formula.FreshnessInfo.volatile.RAND = true; // remember
         break;

      }

   operand.push(result);

   return null;

}

// Add to function list
SocialCalc.Formula.FunctionList["FALSE"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "test"];
SocialCalc.Formula.FunctionList["NA"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "test"];
SocialCalc.Formula.FunctionList["NOW"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "datetime"];
SocialCalc.Formula.FunctionList["RAND"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "math"];
SocialCalc.Formula.FunctionList["PI"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "math"];
SocialCalc.Formula.FunctionList["TODAY"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "datetime"];
SocialCalc.Formula.FunctionList["TRUE"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "test"];

//
// * * * * * FINANCIAL FUNCTIONS * * * * *
//

/*
#
# DDB(cost,salvage,lifetime,period,[method])
#
# Depreciation, method defaults to 2 for double-declining balance
# See: http://en.wikipedia.org/wiki/Depreciation
#
*/

SocialCalc.Formula.DDBFunction = function(fname, operand, foperand, sheet) {

   var method, depreciation, accumulateddepreciation, i;
   var scf = SocialCalc.Formula;

   var cost = scf.OperandAsNumber(sheet, foperand);
   var salvage = scf.OperandAsNumber(sheet, foperand);
   var lifetime = scf.OperandAsNumber(sheet, foperand);
   var period = scf.OperandAsNumber(sheet, foperand);

   if (scf.CheckForErrorValue(operand, cost)) return;
   if (scf.CheckForErrorValue(operand, salvage)) return;
   if (scf.CheckForErrorValue(operand, lifetime)) return;
   if (scf.CheckForErrorValue(operand, period)) return;

   if (lifetime.value < 1) {
      scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncddblife);
      return 0;
      }

   method = {value: 2, type: "n"};
   if (foperand.length > 0 ) {
      method = scf.OperandAsNumber(sheet, foperand);
      }
   if (foperand.length != 0) {
      scf.FunctionArgsError(fname, operand);
      return 0;
      }
   if (scf.CheckForErrorValue(operand, method)) return;

   depreciation = 0; // calculated for each period
   accumulateddepreciation = 0; // accumulated by adding each period's

   for (i=1; i<=period.value-0 && i<=lifetime.value; i++) { // calculate for each period based on net from previous
      depreciation = (cost.value - accumulateddepreciation) * (method.value / lifetime.value);
      if (cost.value - accumulateddepreciation - depreciation < salvage.value) { // don't go lower than salvage value
         depreciation = cost.value - accumulateddepreciation - salvage.value;
         }
      accumulateddepreciation += depreciation;
      }

   scf.PushOperand(operand, 'n$', depreciation);

   return;

   }

SocialCalc.Formula.FunctionList["DDB"] = [SocialCalc.Formula.DDBFunction, -4, "ddb", "", "financial"];

/*
#
# SLN(cost,salvage,lifetime)
#
# Depreciation for each period by straight-line method
# See: http://en.wikipedia.org/wiki/Depreciation
#
*/

SocialCalc.Formula.SLNFunction = function(fname, operand, foperand, sheet) {

   var depreciation;
   var scf = SocialCalc.Formula;

   var cost = scf.OperandAsNumber(sheet, foperand);
   var salvage = scf.OperandAsNumber(sheet, foperand);
   var lifetime = scf.OperandAsNumber(sheet, foperand);

   if (scf.CheckForErrorValue(operand, cost)) return;
   if (scf.CheckForErrorValue(operand, salvage)) return;
   if (scf.CheckForErrorValue(operand, lifetime)) return;

   if (lifetime.value < 1) {
      scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncslnlife);
      return 0;
      }

   depreciation = (cost.value - salvage.value) / lifetime.value;

   scf.PushOperand(operand, 'n$', depreciation);

   return;

   }

SocialCalc.Formula.FunctionList["SLN"] = [SocialCalc.Formula.SLNFunction, 3, "csl", "", "financial"];

/*
#
# SYD(cost,salvage,lifetime,period)
#
# Depreciation by Sum of Year's Digits method
#
*/

SocialCalc.Formula.SYDFunction = function(fname, operand, foperand, sheet) {

   var depreciation, sumperiods;
   var scf = SocialCalc.Formula;

   var cost = scf.OperandAsNumber(sheet, foperand);
   var salvage = scf.OperandAsNumber(sheet, foperand);
   var lifetime = scf.OperandAsNumber(sheet, foperand);
   var period = scf.OperandAsNumber(sheet, foperand);

   if (scf.CheckForErrorValue(operand, cost)) return;
   if (scf.CheckForErrorValue(operand, salvage)) return;
   if (scf.CheckForErrorValue(operand, lifetime)) return;
   if (scf.CheckForErrorValue(operand, period)) return;

   if (lifetime.value < 1 || period.value <= 0) {
      scf.PushOperand(operand, "e#NUM!", 0);
      return 0;
      }

   sumperiods = ((lifetime.value + 1) * lifetime.value)/2; // add up 1 through lifetime
   depreciation = (cost.value - salvage.value) * (lifetime.value - period.value + 1) / sumperiods; // calc depreciation

   scf.PushOperand(operand, 'n$', depreciation);

   return;

   }

SocialCalc.Formula.FunctionList["SYD"] = [SocialCalc.Formula.SYDFunction, 4, "cslp", "", "financial"];

/*
#
# FV(rate, n, payment, [pv, [paytype]])
# NPER(rate, payment, pv, [fv, [paytype]])
# PMT(rate, n, pv, [fv, [paytype]])
# PV(rate, n, payment, [fv, [paytype]])
# RATE(n, payment, pv, [fv, [paytype, [guess]]])
#
# Following the Open Document Format formula specification:
#
#    PV = - Fv - (Payment * Nper) [if rate equals 0]
#    Pv*(1+Rate)^Nper + Payment * (1 + Rate*PaymentType) * ( (1+Rate)^nper -1)/Rate + Fv = 0
#
# For each function, the formulas are solved for the appropriate value (transformed using
# basic algebra).
#
*/

SocialCalc.Formula.InterestFunctions = function(fname, operand, foperand, sheet) {

   var resulttype, result, dval, _eval, fval;
   var pv, fv, rate, n, payment, paytype, guess, part1, part2, part3, part4, part5;
   var olddelta, maxloop, tries, deltaepsilon, rate, oldrate, m;

   var scf = SocialCalc.Formula;

   var aval = scf.OperandAsNumber(sheet, foperand);
   var bval = scf.OperandAsNumber(sheet, foperand);
   var cval = scf.OperandAsNumber(sheet, foperand);

   resulttype = scf.LookupResultType(aval.type, bval.type, scf.TypeLookupTable.twoargnumeric);
   resulttype = scf.LookupResultType(resulttype, cval.type, scf.TypeLookupTable.twoargnumeric);
   if (foperand.length) { // optional arguments
      dval = scf.OperandAsNumber(sheet, foperand);
      resulttype = scf.LookupResultType(resulttype, dval.type, scf.TypeLookupTable.twoargnumeric);
      if (foperand.length) { // optional arguments
         _eval = scf.OperandAsNumber(sheet, foperand);
         resulttype = scf.LookupResultType(resulttype, _eval.type, scf.TypeLookupTable.twoargnumeric);
         if (foperand.length) { // optional arguments
            if (fname != "RATE") { // only rate has 6 possible args
               scf.FunctionArgsError(fname, operand);
               return 0;
               }
            fval = scf.OperandAsNumber(sheet, foperand);
            resulttype = scf.LookupResultType(resulttype, fval.type, scf.TypeLookupTable.twoargnumeric);
            }
         }
      }

   if (resulttype == "n") {
      switch (fname) {
         case "FV": // FV(rate, n, payment, [pv, [paytype]])
            rate = aval.value;
            n = bval.value;
            payment = cval.value;
            pv = dval!=null ? dval.value : 0; // get value if present, or use default
            paytype = _eval!=null ? (_eval.value ? 1 : 0) : 0;
            if (rate == 0) { // simple calculation if no interest
               fv = -pv - (payment * n);
               }
            else {
               fv = -(pv*Math.pow(1+rate,n) + payment * (1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate);
               }
            result = fv;
            resulttype = 'n$';
            break;

         case "NPER": // NPER(rate, payment, pv, [fv, [paytype]])
            rate = aval.value;
            payment = bval.value;
            pv = cval.value;
            fv = dval!=null ? dval.value : 0;
            paytype = _eval!=null ? (_eval.value ? 1 : 0) : 0;
            if (rate == 0) { // simple calculation if no interest
               if (payment == 0) {
                  scf.PushOperand(operand, "e#NUM!", 0);
                  return;
                  }
               n = (pv + fv)/(-payment);
               }
            else {
               part1 = payment * (1 + rate * paytype) / rate;
               part2 = pv + part1;
               if (part2 == 0 || rate <= -1) {
                  scf.PushOperand(operand, "e#NUM!", 0);
                  return;
                  }
               part3 = (part1 - fv) / part2;
               if (part3 <= 0) {
                  scf.PushOperand(operand, "e#NUM!", 0);
                  return;
                  }
               part4 = Math.log(part3);
               part5 = Math.log(1 + rate); // rate > -1
               n = part4/part5;
               }
            result = n;
            resulttype = 'n';
            break;

         case "PMT": // PMT(rate, n, pv, [fv, [paytype]])
            rate = aval.value;
            n = bval.value;
            pv = cval.value;
            fv = dval!=null ? dval.value : 0;
            paytype = _eval!=null ? (_eval.value ? 1 : 0) : 0;
            if (n == 0) {
               scf.PushOperand(operand, "e#NUM!", 0);
               return;
               }
            else if (rate == 0) { // simple calculation if no interest
               payment = (fv - pv)/n;
               }
            else {
               payment = (0 - fv - pv*Math.pow(1+rate,n))/((1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate);
               }
            result = payment;
            resulttype = 'n$';
            break;

         case "PV": // PV(rate, n, payment, [fv, [paytype]])
            rate = aval.value;
            n = bval.value;
            payment = cval.value;
            fv = dval!=null ? dval.value : 0;
            paytype = _eval!=null ? (_eval.value ? 1 : 0) : 0;
            if (rate == -1) {
               scf.PushOperand(operand, "e#DIV/0!", 0);
               return;
               }
            else if (rate == 0) { // simple calculation if no interest
               pv = -fv - (payment * n);
               }
            else {
               pv = (-fv - payment * (1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate)/(Math.pow(1+rate,n));
               }
            result = pv;
            resulttype = 'n$';
            break;

            case "RATE": // RATE(n, payment, pv, [fv, [paytype, [guess]]])
               n = aval.value;
               payment = bval.value;
               pv = cval.value;
               fv = dval!=null ? dval.value : 0;
               paytype = _eval!=null ? (_eval.value ? 1 : 0) : 0;
               guess = fval!=null ? fval.value : 0.1;

               // rate is calculated by repeated approximations
               // The deltas are used to calculate new guesses

               maxloop = 100;
               tries = 0;
               delta = 1;
               epsilon = 0.0000001; // this is close enough
               rate = guess || 0.00000001; // zero is not allowed
               while ((delta >= 0 ? delta : -delta) > epsilon && (rate != oldrate)) {
                  delta = fv + pv*Math.pow(1+rate,n) + payment * (1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate;
                  if (olddelta!=null) {
                     m = (delta - olddelta)/(rate - oldrate) || .001; // get slope (not zero)
                     oldrate = rate;
                     rate = rate - delta / m; // look for zero crossing
                     olddelta = delta;
                     }
                  else { // first time - no old values
                     oldrate = rate;
                     rate = 1.1 * rate;
                     olddelta = delta;
                     }
                  tries++;
                  if (tries >= maxloop) { // didn't converge yet
                     scf.PushOperand(operand, "e#NUM!", 0);
                     return;
                     }
                  }
               result = rate;
               resulttype = 'n%';
               break;
         }
      }
 
   scf.PushOperand(operand, resulttype, result);

   return;

   }

SocialCalc.Formula.FunctionList["FV"] = [SocialCalc.Formula.InterestFunctions, -3, "fv", "", "financial"];
SocialCalc.Formula.FunctionList["NPER"] = [SocialCalc.Formula.InterestFunctions, -3, "nper", "", "financial"];
SocialCalc.Formula.FunctionList["PMT"] = [SocialCalc.Formula.InterestFunctions, -3, "pmt", "", "financial"];
SocialCalc.Formula.FunctionList["PV"] = [SocialCalc.Formula.InterestFunctions, -3, "pv", "", "financial"];
SocialCalc.Formula.FunctionList["RATE"] = [SocialCalc.Formula.InterestFunctions, -3, "rate", "", "financial"];

/*
#
# NPV(rate,v1,v2,c1:c2,...)
#
*/

SocialCalc.Formula.NPVFunction = function(fname, operand, foperand, sheet) {

   var resulttypenpv, rate, sum, factor, value1;

   var scf = SocialCalc.Formula;

   var rate = scf.OperandAsNumber(sheet, foperand);
   if (scf.CheckForErrorValue(operand, rate)) return;

   sum = 0;
   resulttypenpv = "n";
   factor = 1;

   while (foperand.length) {
      value1 = scf.OperandValueAndType(sheet, foperand);
      if (value1.type.charAt(0) == "n") {
         factor *= (1 + rate.value);
         if (factor == 0) {
            scf.PushOperand(operand, "e#DIV/0!", 0);
            return;
            }
         sum += value1.value / factor;
         resulttypenpv = scf.LookupResultType(value1.type, resulttypenpv || value1.type, scf.TypeLookupTable.plus);
         }
      else if (value1.type.charAt(0) == "e" && resulttypenpv.charAt(0) != "e") {
         resulttypenpv = value1.type;
         break;
         }
      }

   if (resulttypenpv.charAt(0) == "n") {
      resulttypenpv = 'n$';
      }

   scf.PushOperand(operand, resulttypenpv, sum);

   return;

   }

SocialCalc.Formula.FunctionList["NPV"] = [SocialCalc.Formula.NPVFunction, -2, "npv", "", "financial"];

/*
#
# IRR(c1:c2,[guess])
#
*/

SocialCalc.Formula.IRRFunction = function(fname, operand, foperand, sheet) {

   var value1, guess, oldsum, maxloop, tries, epsilon, rate, oldrate, m, sum, factor, i;
   var rangeoperand = [];
   var cashflows = [];

   var scf = SocialCalc.Formula;

   rangeoperand.push(foperand.pop()); // first operand is a range

   while (rangeoperand.length) { // get values from range so we can do iterative approximations
      value1 = scf.OperandValueAndType(sheet, rangeoperand);
      if (value1.type.charAt(0) == "n") {
         cashflows.push(value1.value);
         }
      else if (value1.type.charAt(0) == "e") {
         scf.PushOperand(operand, "e#VALUE!", 0);
         return;
         }
      }

   if (!cashflows.length) {
      scf.PushOperand(operand, "e#NUM!", 0);
      return;
      }

   guess = {value: 0};

   if (foperand.length) { // guess is provided
      guess = scf.OperandAsNumber(sheet, foperand);
      if (guess.type.charAt(0) != "n" && guess.type.charAt(0) != "b") {
         scf.PushOperand(operand, "e#VALUE!", 0);
         return;
         }
      if (foperand.length) { // should be no more args
         scf.FunctionArgsError(fname, operand);
         return;
         }
      }

   guess.value = guess.value || 0.1;

   // rate is calculated by repeated approximations
   // The deltas are used to calculate new guesses

   maxloop = 20;
   tries = 0;
   epsilon = 0.0000001; // this is close enough
   rate = guess.value;
   sum = 1;

   while ((sum >= 0 ? sum : -sum) > epsilon && (rate != oldrate)) {
      sum = 0;
      factor = 1;
      for (i=0; i<cashflows.length; i++) {
         factor *= (1 + rate);
         if (factor == 0) {
            scf.PushOperand(operand, "e#DIV/0!", 0);
            return;
            }
         sum += cashflows[i] / factor;
         }

      if (oldsum!=null) {
         m = (sum - oldsum)/(rate - oldrate); // get slope
         oldrate = rate;
         rate = rate - sum / m; // look for zero crossing
         oldsum = sum;
         }
      else { // first time - no old values
         oldrate = rate;
         rate = 1.1 * rate;
         oldsum = sum;
         }
      tries++;
      if (tries >= maxloop) { // didn't converge yet
         scf.PushOperand(operand, "e#NUM!", 0);
         return;
         }
      }

   scf.PushOperand(operand, 'n%', rate);

   return;

   }

SocialCalc.Formula.FunctionList["IRR"] = [SocialCalc.Formula.IRRFunction, -1, "irr", "", "financial"];

//
// SHEET CACHE
//

SocialCalc.Formula.SheetCache = {

   // Sheet data: Attributes are each sheet in the cache with values of an object with:
   //
   //    sheet: sheet-obj (or null, meaning not found)
   //    recalcstate: constants.asloaded = as loaded
   //                 constants.recalcing = being recalced now
   //                 constants.recalcdone = recalc done
   //    name: name of sheet (in case just have object and don't know name)
   //

   sheets: {},

   // Waiting for loading:
   // If sheet is not in cache, this is set to the sheetname being loaded
   // so it can be tested in the recalc loop to start load and then wait until restarted.
   // Reset to null before restarting.

   waitingForLoading: null,

   // Constants to use for setting sheets[*].recalcstate:

   constants: {asloaded: 0, recalcing: 1, recalcdone: 2},

   loadsheet: null // (deprecated - use SocialCalc.RecalcInfo.LoadSheet)

   };

//
// othersheet = SocialCalc.Formula.FindInSheetCache(sheetname)
//
// Returns a SocialCalc.Sheet object corresponding to string sheetname
// or null if the sheet is not available or in error.
//
// Each sheet is loaded only once and then stored in a cache.
// Loading is handled elsewhere, e.g., in the recalc loop.
//

SocialCalc.Formula.FindInSheetCache = function(sheetname) {

   var str;
   var sfsc = SocialCalc.Formula.SheetCache;

   var nsheetname = SocialCalc.Formula.NormalizeSheetName(sheetname); // normalize different versions

   if (sfsc.sheets[nsheetname]) { // a sheet by that name is in the cache already
      return sfsc.sheets[nsheetname].sheet; // return it
      }

   if (sfsc.waitingForLoading) { // waiting already - only queue up one
      return null; // return not found
      }

   if (sfsc.loadsheet) { // Deprecated old format synchronous callback
alert("Using SocialCalc.Formula.SheetCache.loadsheet - deprecated");
      return SocialCalc.Formula.AddSheetToCache(nsheetname, sfsc.loadsheet(nsheetname));
      }

   sfsc.waitingForLoading = nsheetname; // let recalc loop know that we have a sheet to load

   return null; // return not found

   }

//
// newsheet = SocialCalc.Formula.AddSheetToCache(sheetname, str, live)
//
// Adds a new sheet to the sheet cache.
// Returns the sheet object filled out with the str (a saved sheet).
//

SocialCalc.Formula.AddSheetToCache = function(sheetname, str, live) {

   var newsheet = null;
   var sfsc = SocialCalc.Formula.SheetCache;
   var sfscc = sfsc.constants;
   var newsheetname = SocialCalc.Formula.NormalizeSheetName(sheetname);

   if (str) {
      newsheet = new SocialCalc.Sheet();
      newsheet.ParseSheetSave(str);
      }

   sfsc.sheets[newsheetname] = {sheet: newsheet, recalcstate: sfscc.asloaded, name: newsheetname};

   SocialCalc.Formula.FreshnessInfo.sheets[newsheetname] = (typeof(live) == "undefined" || live === false);

   return newsheet;

   }

//
// nsheet = SocialCalc.Formula.NormalizeSheetName(sheetname)
//

SocialCalc.Formula.NormalizeSheetName = function(sheetname) {

   if (SocialCalc.Callbacks.NormalizeSheetName) {
      return SocialCalc.Callbacks.NormalizeSheetName(sheetname);
      }
   else {
      return sheetname.toLowerCase();
      }
   }

//
// REMOTE FUNCTION INFO
//

SocialCalc.Formula.RemoteFunctionInfo = {

   // Waiting for server:
   // If waiting for an XHR response from the server, this is set to some non-blank status text
   // so it can be tested in the recalc loop to start load and then wait until restarted.
   // Reset to null before restarting.

   waitingForServer: null

   };

//
// FRESHNESS INFO
//
// This information is generated during recalc.
// It may be used to help determine when the recalc data in a spreadsheet
// may be out of date.
// For example, it may be used to display a message like:
// "Dependent on sheet 'FOO' which was updated more recently than this printout"

SocialCalc.Formula.FreshnessInfo = {

   // For each external sheet referenced successfully an attribute of that name with value true to keep the sheet cached.
   // Value false means the sheet is reloaded at each recalc.

   sheets: {},

   // For each volatile function that is called an attribute of that name with value true.

   volatile: {},

   // Set to false when started and true when recalc completes

   recalc_completed: false

   };

SocialCalc.Formula.FreshnessInfoReset = function() {

   var scffi = SocialCalc.Formula.FreshnessInfo;
   var scfsc = SocialCalc.Formula.SheetCache;

   // Loop through sheets freshness, deleting cached sheets that should be reloaded.

   for (var sheet in scffi.sheets) {
      if (scffi.sheets[sheet] === false) {
         delete scfsc.sheets[sheet];
         }
      }
   
   // Reset freshness info.

   scffi.sheets = {};
   scffi.volatile = {};
   scffi.recalc_completed = false;

   }

//
// MISC ROUTINES
//

//
// result = SocialCalc.Formula.PlainCoord(coord)
//
// Returns: coord without any $'s
//

SocialCalc.Formula.PlainCoord = function(coord) {

   if (coord.indexOf("$") == -1) return coord;

   return coord.replace(/\$/g, ""); // remove any $'s

   }

//
// result = SocialCalc.Formula.OrderRangeParts(coord1, coord2)
//
// Returns: {c1: col, r1: row, c2: col, r2 = row} with c1/r1 upper left
//

SocialCalc.Formula.OrderRangeParts = function(coord1, coord2) {

   var cr1, cr2;
   var result = {};

   cr1 = SocialCalc.coordToCr(coord1);
   cr2 = SocialCalc.coordToCr(coord2);
   if (cr1.col > cr2.col) { result.c1 = cr2.col; result.c2 = cr1.col; }
   else { result.c1 = cr1.col; result.c2 = cr2.col; }
   if (cr1.row > cr2.row) { result.r1 = cr2.row; result.r2 = cr1.row; }
   else { result.r1 = cr1.row; result.r2 = cr2.row; }

   return result;

   }

//
// cond = SocialCalc.Formula.TestCriteria(value, type, criteria)
//
// Determines whether a value/type meets the criteria.
// A criteria can be a numeric value, text beginning with <, <=, =, >=, >, <>, text by itself is start of text to match.
// Used by a variety of functions, including the "D" functions (DSUM, etc.).
//
// Returns true or false
//

SocialCalc.Formula.TestCriteria = function(value, type, criteria) {

   var comparitor, basestring, basevalue, cond, testvalue;

   if (criteria == null) { // undefined (e.g., error value) is always false
      return false;
      }

   criteria = criteria + "";
   comparitor = criteria.charAt(0); // look for comparitor
   if (comparitor == "=" || comparitor == "<" || comparitor == ">") {
      basestring = criteria.substring(1);
      }
   else {
      comparitor = criteria.substring(0,2);
      if (comparitor == "<=" || comparitor == "<>" || comparitor == ">=") {
         basestring = criteria.substring(2);
         }
      else {
         comparitor = "none";
         basestring = criteria;
         }
      }

   basevalue = SocialCalc.DetermineValueType(basestring); // get type of value being compared
   if (!basevalue.type) { // no criteria base value given
      if (comparitor == "none") { // blank criteria matches nothing
         return false;
         }
      if (type.charAt(0) == "b") { // comparing to empty cell
         if (comparitor == "=") { // empty equals empty
            return true;
            }
         }
      else {
         if (comparitor == "<>") { // "something" does not equal empty
            return true;
            }
         }
      return false; // otherwise false
      }

   cond = false;

   if (basevalue.type.charAt(0) == "n" && type.charAt(0) == "t") { // criteria is number, but value is text
      testvalue = SocialCalc.DetermineValueType(value);
      if (testvalue.type.charAt(0) == "n") { // could be number - make it one
         value = testvalue.value;
         type = testvalue.type;
         }
      }

   if (type.charAt(0) == "n" && basevalue.type.charAt(0) == "n") { // compare two numbers
      value = value - 0; // make sure numbers
      basevalue.value = basevalue.value - 0;
      switch (comparitor) {
         case "<":
            cond = value < basevalue.value;
            break;

         case "<=":
            cond = value <= basevalue.value;
            break;

         case "=":
         case "none":
            cond = value == basevalue.value;
            break;

         case ">=":
            cond = value >= basevalue.value;
            break;

         case ">":
            cond = value > basevalue.value;
            break;

         case "<>":
            cond = value != basevalue.value;
            break;
         }
      }

   else if (type.charAt(0) == "e") { // error on left
      cond = false;
      }

   else if (basevalue.type.charAt(0) == "e") { // error on right
      cond = false;
      }

   else { // text, maybe mixed with number or blank
      if (type.charAt(0) == "n") {
         value = SocialCalc.format_number_for_display(value, "n", "");
         }
      if (basevalue.type.charAt(0) == "n") {
         return false; // if number and didn't match already, isn't a match
         }

      value = value ? value.toLowerCase() : "";
      basevalue.value = basevalue.value ? basevalue.value.toLowerCase() : "";

      switch (comparitor) {
         case "<":
            cond = value < basevalue.value;
            break;

         case "<=":
            cond = value <= basevalue.value;
            break;

         case "=":
            cond = value == basevalue.value;
            break;

         case "none":
            cond = value.substring(0, basevalue.value.length) == basevalue.value;
            break;

         case ">=":
            cond = value >= basevalue.value;
            break;

         case ">":
            cond = value > basevalue.value;
            break;

         case "<>":
            cond = value != basevalue.value;
            break;
         }
      }

   return cond;

   }
//
/*
// SocialCalc Number Formatting Library
//
// Part of the SocialCalc package.
//
// (c) Copyright 2008 Socialtext, Inc.
// All Rights Reserved.
//
// The contents of this file are subject to the Artistic License 2.0; you may not
// use this file except in compliance with the License. You may obtain a copy of 
// the License at http://socialcalc.org/licenses/al-20/.
//
// Some of the other files in the SocialCalc package are licensed under
// different licenses. Please note the licenses of the modules you use.
//
// Code History:
//
// Initially coded by Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
// Based in part on the SocialCalc 1.1.0 code written in Perl.
// The SocialCalc 1.1.0 code was:
//    Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc.
//    All Rights Reserved.
//    Portions (c) Copyright 2007 Socialtext, Inc.
//    All Rights Reserved.
// The Perl SocialCalc started as modifications to the wikiCalc(R) program, version 1.0.
// wikiCalc 1.0 was written by Software Garden, Inc.
// Unless otherwise specified, referring to "SocialCalc" in comments refers to this
// JavaScript version of the code, not the SocialCalc Perl code.
//
*/

   var SocialCalc;
   if (!SocialCalc) SocialCalc = {}; // May be used with other SocialCalc libraries or standalone

SocialCalc.FormatNumber = {};

SocialCalc.FormatNumber.format_definitions = {}; // Parsed formats are stored here globally

// Most constants that are often customized for localization are in the SocialCalc.Constants module.
// If you use this module standalone, provide at least the "FormatNumber" values.
//

// The following values may be customized externally for further localization of the format definitions themselves,
// but that would make them incompatible with other uses and is discouraged.
//

SocialCalc.FormatNumber.separatorchar = ",";
SocialCalc.FormatNumber.decimalchar = ".";
SocialCalc.FormatNumber.daynames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
SocialCalc.FormatNumber.daynames3 = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
SocialCalc.FormatNumber.monthnames3 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
SocialCalc.FormatNumber.monthnames = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                                      "October", "November", "December"];

SocialCalc.FormatNumber.allowedcolors =
   {BLACK: "#000000", BLUE: "#0000FF", CYAN: "#00FFFF", GREEN: "#00FF00", MAGENTA: "#FF00FF",
    RED: "#FF0000", WHITE: "#FFFFFF", YELLOW: "#FFFF00"};

SocialCalc.FormatNumber.alloweddates =
   {H: "h]", M: "m]", MM: "mm]", S: "s]", SS: "ss]"};

// Other constants

SocialCalc.FormatNumber.commands =
   {copy: 1, color: 2, integer_placeholder: 3, fraction_placeholder: 4, decimal: 5,
    currency: 6, general:7, separator: 8, date: 9, comparison: 10, section: 11, style: 12};

SocialCalc.FormatNumber.datevalues = {julian_offset: 2415019, seconds_in_a_day: 24 * 60 * 60, seconds_in_an_hour: 60 * 60};

/* *******************

 result = SocialCalc.FormatNumber.formatNumberWithFormat = function(rawvalue, format_string, currency_char)

************************* */

SocialCalc.FormatNumber.formatNumberWithFormat = function(rawvalue, format_string, currency_char) {

   var scc = SocialCalc.Constants;
   var scfn = SocialCalc.FormatNumber;

   var op, operandstr, fromend, cval, operandstrlc;
   var startval, estartval;
   var hrs, mins, secs, ehrs, emins, esecs, ampmstr, ymd;
   var minOK, mpos;
   var result="";
   var thisformat;
   var section, gotcomparison, compop, compval, cpos, oppos;
   var sectioninfo;
   var i, decimalscale, scaledvalue, strvalue, strparts, integervalue, fractionvalue;
   var integerdigits2, integerpos, fractionpos, textcolor, textstyle, separatorchar, decimalchar;
   var value; // working copy to change sign, etc.

   if (typeof(rawvalue) == "string" && !rawvalue.length) return "";

   value = rawvalue-0; // make sure a number
   if (!isFinite(value)) {
      if (typeof(rawvalue) == "string") { // if original was a string, try to format it
         return scfn.formatTextWithFormat(rawvalue, format_string);
         }
      else {
         return "NaN";
         }
      }
   rawvalue = value;

   var negativevalue = value < 0 ? 1 : 0; // determine sign, etc.
   if (negativevalue) value = -value;
   var zerovalue = value == 0 ? 1 : 0;

   currency_char = currency_char || scc.FormatNumber_DefaultCurrency;

   scfn.parse_format_string(scfn.format_definitions, format_string); // make sure format is parsed
   thisformat = scfn.format_definitions[format_string]; // Get format structure

   if (!thisformat) throw "Format not parsed error!";

   section = thisformat.sectioninfo.length - 1; // get number of sections - 1

   if (thisformat.hascomparison) { // has comparisons - determine which section
      section = 0; // set to which section we will use
      gotcomparison = 0; // this section has no comparison
      for (cpos=0; ;cpos++) { // scan for comparisons
         op = thisformat.operators[cpos];
         operandstr = thisformat.operands[cpos]; // get next operator and operand
         if (!op) { // at end with no match
            if (gotcomparison) { // if comparison but no match
               format_string = "General"; // use default of General
               scfn.parse_format_string(scfn.format_definitions, format_string);
               thisformat = scfn.format_definitions[format_string];
               section = 0;
               }
            break; // if no comparision, matches on this section
            }
         if (op == scfn.commands.section) { // end of section
            if (!gotcomparison) { // no comparison, so it's a match
               break;
               }
            gotcomparison = 0;
            section++; // check out next one
            continue;
            }
         if (op == scfn.commands.comparison) { // found a comparison - do we meet it?
            i=operandstr.indexOf(":");
            compop=operandstr.substring(0,i);
            compval=operandstr.substring(i+1)-0;
            if ((compop == "<" && rawvalue < compval) ||
                (compop == "<=" && rawvalue <= compval) ||
                (compop == "=" && rawvalue == compval) ||
                (compop == "<>" && rawvalue != compval) ||
                (compop == ">=" && rawvalue >= compval) ||
                (compop == ">" && rawvalue > compval)) { // a match
               break;
               }
            gotcomparison = 1;
            }
         }
      }
   else if (section > 0) { // more than one section (separated by ";")
      if (section == 1) { // two sections
         if (negativevalue) {
            negativevalue = 0; // sign will provided by section, not automatically
            section = 1; // use second section for negative values
            }
         else {
            section = 0; // use first for all others
            }
         }
      else if (section == 2 || section == 3) { // three or four sections
         if (negativevalue) {
            negativevalue = 0; // sign will provided by section, not automatically
            section = 1; // use second section for negative values
            }
         else if (zerovalue) {
            section = 2; // use third section for zero values
            }
         else {
            section = 0; // use first for positive
            }
         }
      }

   sectioninfo = thisformat.sectioninfo[section]; // look at values for our section

   if (sectioninfo.commas > 0) { // scale by thousands
      for (i=0; i<sectioninfo.commas; i++) {
         value /= 1000;
         }
      }
   if (sectioninfo.percent > 0) { // do percent scaling
      for (i=0; i<sectioninfo.percent; i++) {
         value *= 100;
         }
      }

   decimalscale = 1; // cut down to required number of decimal digits
   for (i=0; i<sectioninfo.fractiondigits; i++) {
      decimalscale *= 10;
      }
   scaledvalue = Math.floor(value * decimalscale + 0.5);
   scaledvalue = scaledvalue / decimalscale;

   if (typeof scaledvalue != "number") return "NaN";
   if (!isFinite(scaledvalue)) return "NaN";

   strvalue = scaledvalue+""; // convert to string (Number.toFixed doesn't do all we need)

//   strvalue = value.toFixed(sectioninfo.fractiondigits); // cut down to required number of decimal digits
                                                         // and convert to string

   if (scaledvalue == 0 && (sectioninfo.fractiondigits || sectioninfo.integerdigits)) {
      negativevalue = 0; // no "-0" unless using multiple sections or General
      }

   if (strvalue.indexOf("e")>=0) { // converted to scientific notation
      return rawvalue+""; // Just return plain converted raw value
      }

   strparts=strvalue.match(/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/); // get integer and fraction parts
   if (!strparts) return "NaN"; // if not a number
   integervalue = strparts[1];
   if (!integervalue || integervalue=="0") integervalue="";
   fractionvalue = strparts[2];
   if (!fractionvalue) fractionvalue = "";

   if (sectioninfo.hasdate) { // there are date placeholders
      if (rawvalue < 0) { // bad date
         return "??-???-??&nbsp;??:??:??";
         }
      startval = (rawvalue-Math.floor(rawvalue)) * scfn.datevalues.seconds_in_a_day; // get date/time parts
      estartval = rawvalue * scfn.datevalues.seconds_in_a_day; // do elapsed time version, too
      hrs = Math.floor(startval / scfn.datevalues.seconds_in_an_hour);
      ehrs = Math.floor(estartval / scfn.datevalues.seconds_in_an_hour);
      startval = startval - hrs * scfn.datevalues.seconds_in_an_hour;
      mins = Math.floor(startval / 60);
      emins = Math.floor(estartval / 60);
      secs = startval - mins * 60;
      decimalscale = 1; // round appropriately depending if there is ss.0
      for (i=0; i<sectioninfo.fractiondigits; i++) {
         decimalscale *= 10;
         }
      secs = Math.floor(secs * decimalscale + 0.5);
      secs = secs / decimalscale;
      esecs = Math.floor(estartval * decimalscale + 0.5);
      esecs = esecs / decimalscale;
      if (secs >= 60) { // handle round up into next second, minute, etc.
         secs = 0;
         mins++; emins++;
         if (mins >= 60) {
            mins = 0;
            hrs++; ehrs++;
            if (hrs >= 24) {
               hrs = 0;
               rawvalue++;
               }
            }
         }
      fractionvalue = (secs-Math.floor(secs))+""; // for "hh:mm:ss.000"
      fractionvalue = fractionvalue.substring(2); // skip "0."

      ymd = SocialCalc.FormatNumber.convert_date_julian_to_gregorian(Math.floor(rawvalue+scfn.datevalues.julian_offset));

      minOK = 0; // says "m" can be minutes if true
      mspos = sectioninfo.sectionstart; // m scan position in ops
      for ( ; ; mspos++) { // scan for "m" and "mm" to see if any minutes fields, and am/pm
         op = thisformat.operators[mspos];
         operandstr = thisformat.operands[mspos]; // get next operator and operand
         if (!op) break; // don't go past end
         if (op==scfn.commands.section) break;
         if (op==scfn.commands.date) {
            if ((operandstr.toLowerCase()=="am/pm" || operandstr.toLowerCase()=="a/p") && !ampmstr) {
               if (hrs >= 12) {
                  hrs -= 12;
                  ampmstr = operandstr.toLowerCase()=="a/p" ? scc.s_FormatNumber_pm1 : scc.s_FormatNumber_pm; // "P" : "PM";
                  }
               else {
                  ampmstr = operandstr.toLowerCase()=="a/p" ? scc.s_FormatNumber_am1 : scc.s_FormatNumber_am; // "A" : "AM";
                  }
               if (operandstr.indexOf(ampmstr)<0)
                  ampmstr = ampmstr.toLowerCase(); // have case match case in format
               }
            if (minOK && (operandstr=="m" || operandstr=="mm")) {
               thisformat.operands[mspos] += "in"; // turn into "min" or "mmin"
               }
            if (operandstr.charAt(0)=="h") {
               minOK = 1; // m following h or hh or [h] is minutes not months
               }
            else {
               minOK = 0;
               }
            }
         else if (op!=scfn.commands.copy) { // copying chars can be between h and m
            minOK = 0;
            }
         }
      minOK = 0;
      for (--mspos; ; mspos--) { // scan other way for s after m
         op = thisformat.operators[mspos];
         operandstr = thisformat.operands[mspos]; // get next operator and operand
         if (!op) break; // don't go past end
         if (op==scfn.commands.section) break;
         if (op==scfn.commands.date) {
            if (minOK && (operandstr=="m" || operandstr=="mm")) {
               thisformat.operands[mspos] += "in"; // turn into "min" or "mmin"
               }
            if (operandstr=="ss") {
               minOK = 1; // m before ss is minutes not months
               }
            else {
               minOK = 0;
               }
            }
         else if (op!=scfn.commands.copy) { // copying chars can be between ss and m
            minOK = 0;
            }
         }
      }

   integerdigits2 = 0; // init counters, etc.
   integerpos = 0;
   fractionpos = 0;
   textcolor = "";
   textstyle = "";
   separatorchar = scc.FormatNumber_separatorchar;
   if (separatorchar.indexOf(" ")>=0) separatorchar = separatorchar.replace(/ /g, "&nbsp;");
   decimalchar = scc.FormatNumber_decimalchar;
   if (decimalchar.indexOf(" ")>=0) decimalchar = decimalchar.replace(/ /g, "&nbsp;");

   oppos = sectioninfo.sectionstart;

   while (op = thisformat.operators[oppos]) { // execute format
      operandstr = thisformat.operands[oppos++]; // get next operator and operand

      if (op == scfn.commands.copy) { // put char in result
         result += operandstr;
         }

      else if (op == scfn.commands.color) { // set color
         textcolor = operandstr;
         }

      else if (op == scfn.commands.style) { // set style
         textstyle = operandstr;
         }

      else if (op == scfn.commands.integer_placeholder) { // insert number part
         if (negativevalue) {
            result += "-";
            negativevalue = 0;
            }
         integerdigits2++;
         if (integerdigits2 == 1) { // first one
            if (integervalue.length > sectioninfo.integerdigits) { // see if integer wider than field
               for (;integerpos < (integervalue.length - sectioninfo.integerdigits); integerpos++) {
                  result += integervalue.charAt(integerpos);
                  if (sectioninfo.thousandssep) { // see if this is a separator position
                     fromend = integervalue.length - integerpos - 1;
                     if (fromend > 2 && fromend % 3 == 0) {
                        result += separatorchar;
                        }
                     }
                  }
               }
            }
         if (integervalue.length < sectioninfo.integerdigits
             && integerdigits2 <= sectioninfo.integerdigits - integervalue.length) { // field is wider than value
            if (operandstr == "0" || operandstr == "?") { // fill with appropriate characters
               result += operandstr == "0" ? "0" : "&nbsp;";
               if (sectioninfo.thousandssep) { // see if this is a separator position
                  fromend = sectioninfo.integerdigits - integerdigits2;
                  if (fromend > 2 && fromend % 3 == 0) {
                     result += separatorchar;
                     }
                  }
               }
            }
         else { // normal integer digit - add it
            result += integervalue.charAt(integerpos);
            if (sectioninfo.thousandssep) { // see if this is a separator position
               fromend = integervalue.length - integerpos - 1;
               if (fromend > 2 && fromend % 3 == 0) {
                  result += separatorchar;
                  }
               }
            integerpos++;
            }
         }
      else if (op == scfn.commands.fraction_placeholder) { // add fraction part of number
         if (fractionpos >= fractionvalue.length) {
            if (operandstr == "0" || operandstr == "?") {
               result += operandstr == "0" ? "0" : "&nbsp;";
               }
            }
         else {
            result += fractionvalue.charAt(fractionpos);
            }
         fractionpos++;
         }

      else if (op == scfn.commands.decimal) { // decimal point
         if (negativevalue) {
            result += "-";
            negativevalue = 0;
            }
         result += decimalchar;
         }

      else if (op == scfn.commands.currency) { // currency symbol
         if (negativevalue) {
            result += "-";
            negativevalue = 0;
            }
         result += operandstr;
         }

      else if (op == scfn.commands.general) { // insert "General" conversion

         // *** Cut down number of significant digits to avoid floating point artifacts:

         if (value!=0) { // only if non-zero
            var factor = Math.floor(Math.LOG10E * Math.log(value)); // get integer magnitude as a power of 10
            factor = Math.pow(10, 13-factor); // turn into scaling factor
            value = Math.floor(factor * value + 0.5)/factor; // scale positive value, round, undo scaling
            if (!isFinite(value)) return "NaN";
            }
         if (negativevalue) {
            result += "-";
            }
         strvalue = value+""; // convert original value to string
         if (strvalue.indexOf("e")>=0) { // converted to scientific notation
            result += strvalue;
            continue;
            }
         strparts=strvalue.match(/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/); // get integer and fraction parts
         integervalue = strparts[1];
         if (!integervalue || integervalue=="0") integervalue="";
         fractionvalue = strparts[2];
         if (!fractionvalue) fractionvalue = "";
         integerpos = 0;
         fractionpos = 0;
         if (integervalue.length) {
            for (;integerpos < integervalue.length; integerpos++) {
               result += integervalue.charAt(integerpos);
               if (sectioninfo.thousandssep) { // see if this is a separator position
                  fromend = integervalue.length - integerpos - 1;
                  if (fromend > 2 && fromend % 3 == 0) {
                     result += separatorchar;
                     }
                  }
               }
             }
         else {
            result += "0";
            }
         if (fractionvalue.length) {
            result += decimalchar;
            for (;fractionpos < fractionvalue.length; fractionpos++) {
               result += fractionvalue.charAt(fractionpos);
               }
            }
         }
      else if (op==scfn.commands.date) { // date placeholder
         operandstrlc = operandstr.toLowerCase();
         if (operandstrlc=="y" || operandstrlc=="yy") {
            result += (ymd.year+"").substring(2);
            }
         else if (operandstrlc=="yyyy") {
            result += ymd.year+"";
            }
         else if (operandstrlc=="d") {
            result += ymd.day+"";
            }
         else if (operandstrlc=="dd") {
            cval = 1000 + ymd.day;
            result += (cval+"").substr(2);
            }
         else if (operandstrlc=="ddd") {
            cval = Math.floor(rawvalue+6) % 7;
            result += scc.s_FormatNumber_daynames3[cval];
            }
         else if (operandstrlc=="dddd") {
            cval = Math.floor(rawvalue+6) % 7;
            result += scc.s_FormatNumber_daynames[cval];
            }
         else if (operandstrlc=="m") {
            result += ymd.month+"";
            }
         else if (operandstrlc=="mm") {
            cval = 1000 + ymd.month;
            result += (cval+"").substr(2);
            }
         else if (operandstrlc=="mmm") {
            result += scc.s_FormatNumber_monthnames3[ymd.month-1];
            }
         else if (operandstrlc=="mmmm") {
            result += scc.s_FormatNumber_monthnames[ymd.month-1];
            }
         else if (operandstrlc=="mmmmm") {
            result += scc.s_FormatNumber_monthnames[ymd.month-1].charAt(0);
            }
         else if (operandstrlc=="h") {
            result += hrs+"";
            }
         else if (operandstrlc=="h]") {
            result += ehrs+"";
            }
         else if (operandstrlc=="mmin") {
            cval = (1000 + mins)+"";
            result += cval.substr(2);
            }
         else if (operandstrlc=="mm]") {
            if (emins < 100) {
               cval = (1000 + emins)+"";
               result += cval.substr(2);
               }
            else {
               result += emins+"";
               }
            }
         else if (operandstrlc=="min") {
            result += mins+"";
            }
         else if (operandstrlc=="m]") {
            result += emins+"";
            }
         else if (operandstrlc=="hh") {
            cval = (1000 + hrs)+"";
            result += cval.substr(2);
            }
         else if (operandstrlc=="s") {
            cval = Math.floor(secs);
            result += cval+"";
            }
         else if (operandstrlc=="ss") {
            cval = (1000 + Math.floor(secs))+"";
            result += cval.substr(2);
            }
         else if (operandstrlc=="am/pm" || operandstrlc=="a/p") {
            result += ampmstr;
            }
         else if (operandstrlc=="ss]") {
            if (esecs < 100) {
               cval = (1000 + Math.floor(esecs))+"";
               result += cval.substr(2);
               }
            else {
               cval = Math.floor(esecs);
               result += cval+"";
               }
            }
         }
      else if (op == scfn.commands.section) { // end of section
         break;
         }

      else if (op == scfn.commands.comparison) { // ignore
         continue;
         }

      else {
         result += "!! Parse error !!";
         }
      }

   if (textcolor) {
      result = '<span style="color:'+textcolor+';">'+result+'</span>';
      }
   if (textstyle) {
      result = '<span style="'+textstyle+';">'+result+'</span>';
      }

   return result;

   };

/* *******************

 result = SocialCalc.FormatNumber.formatTextWithFormat = function(rawvalue, format_string)

************************* */

SocialCalc.FormatNumber.formatTextWithFormat = function(rawvalue, format_string) {

   var scc = SocialCalc.Constants;
   var scfn = SocialCalc.FormatNumber;
   var value = rawvalue+"";
   var result = "";
   var section;
   var sectioninfo;
   var oppos;
   var operandstr;
   var textcolor = "";
   var textstyle = "";

   scfn.parse_format_string(scfn.format_definitions, format_string); // make sure format is parsed
   thisformat = scfn.format_definitions[format_string]; // Get format structure

   if (!thisformat) throw "Format not parsed error!";

   section = thisformat.sectioninfo.length - 1; // get number of sections - 1
   if (section == 0) {
      section = 0;
      }
   else if (section == 3) {
      section = 3;
      }
   else {
      return value;
      }

   sectioninfo = thisformat.sectioninfo[section]; // look at values for our section
   oppos = sectioninfo.sectionstart;

   while (op = thisformat.operators[oppos]) { // execute format
      operandstr = thisformat.operands[oppos++]; // get next operator and operand

      if (op == scfn.commands.copy) { // put char in result
         if (operandstr == "@") {
            result += value;
            }
         else {
            result += operandstr.replace(/ /g, "&nbsp;");
            }
         }

      else if (op == scfn.commands.color) { // set color
         textcolor = operandstr;
         }

      else if (op == scfn.commands.style) { // set style
         textstyle = operandstr;
         }
      }

   if (textcolor) {
      result = '<span style="color:'+textcolor+';">'+result+'</span>';
      }
   if (textstyle) {
      result = '<span style="'+textstyle+';">'+result+'</span>';
      }

   return result;

   };

/* *******************

 SocialCalc.FormatNumber.parse_format_string(format_defs, format_string)

 Takes a format string (e.g., "#,##0.00_);(#,##0.00)") and fills in format_defs with the parsed info

 format_defs
    ["#,##0.0"]->{} - elements in the hash are one hash for each format
       .operators->[] - array of operators from parsing the format string (each a number)
       .operands->[] - array of corresponding operators (each usually a string)
       .sectioninfo->[] - one hash for each section of the format
          .start
          .integerdigits
          .fractiondigits
          .commas
          .percent
          .thousandssep
          .hasdates
       .hascomparison - true if any section has [<100], etc.

************************* */

SocialCalc.FormatNumber.parse_format_string = function(format_defs, format_string) {

   var scfn = SocialCalc.FormatNumber;

   var thisformat, section, sectioninfo;
   var integerpart = 1; // start out in integer part
   var lastwasinteger; // last char was an integer placeholder
   var lastwasslash; // last char was a backslash - escaping following character
   var lastwasasterisk; // repeat next char
   var lastwasunderscore; // last char was _ which picks up following char for width
   var inquote, quotestr; // processing a quoted string
   var inbracket, bracketstr, bracketdata; // processing a bracketed string
   var ingeneral, gpos; // checks for characters "General"
   var ampmstr, part; // checks for characters "A/P" and "AM/PM"
   var indate; // keeps track of date/time placeholders
   var chpos; // character position being looked at
   var ch; // character being looked at

   if (format_defs[format_string]) return; // already defined - nothing to do

   thisformat = {operators: [], operands: [], sectioninfo: [{}]}; // create info structure for this format
   format_defs[format_string] = thisformat; // add to other format definitions

   section = 0; // start with section 0
   sectioninfo = thisformat.sectioninfo[section]; // get reference to info for current section
   sectioninfo.sectionstart = 0; // position in operands that starts this section
   sectioninfo.integerdigits = 0; // number of integer-part placeholders
   sectioninfo.fractiondigits = 0; // fraction placeholders
   sectioninfo.commas = 0; // commas encountered, to handle scaling
   sectioninfo.percent = 0; // times to scale by 100

   for (chpos=0; chpos<format_string.length; chpos++) { // parse
      ch = format_string.charAt(chpos); // get next char to examine
      if (inquote) {
         if (ch == '"') {
            inquote = 0;
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(quotestr);
            continue;
            }
         quotestr += ch;
         continue;
         }
      if (inbracket) {
         if (ch==']') {
            inbracket = 0;
            bracketdata=SocialCalc.FormatNumber.parse_format_bracket(bracketstr);
            if (bracketdata.operator==scfn.commands.separator) {
               sectioninfo.thousandssep = 1; // explicit [,]
               continue;
               }
            if (bracketdata.operator==scfn.commands.date) {
               sectioninfo.hasdate = 1;
               }
            if (bracketdata.operator==scfn.commands.comparison) {
               thisformat.hascomparison = 1;
               }
            thisformat.operators.push(bracketdata.operator);
            thisformat.operands.push(bracketdata.operand);
            continue;
            }
         bracketstr += ch;
         continue;
         }
      if (lastwasslash) {
         thisformat.operators.push(scfn.commands.copy);
         thisformat.operands.push(ch);
         lastwasslash=false;
         continue;
         }
      if (lastwasasterisk) {
         thisformat.operators.push(scfn.commands.copy);
         thisformat.operands.push(ch+ch+ch+ch+ch); // do 5 of them since no real tabs
         lastwasasterisk=false;
         continue;
         }
      if (lastwasunderscore) {
         thisformat.operators.push(scfn.commands.copy);
         thisformat.operands.push("&nbsp;");
         lastwasunderscore=false;
         continue;
         }
      if (ingeneral) {
         if ("general".charAt(ingeneral)==ch.toLowerCase()) {
            ingeneral++;
            if (ingeneral == 7) {
               thisformat.operators.push(scfn.commands.general);
               thisformat.operands.push(ch);
               ingeneral=0;
               }
            continue;
            }
         ingeneral = 0;
         }
      if (indate) { // last char was part of a date placeholder
         if (indate.charAt(0)==ch) { // another of the same char
            indate += ch; // accumulate it
            continue;
            }
         thisformat.operators.push(scfn.commands.date); // something else, save date info
         thisformat.operands.push(indate);
         sectioninfo.hasdate=1;
         indate = "";
         }
      if (ampmstr) {
         ampmstr += ch;
         part=ampmstr.toLowerCase();
         if (part!="am/pm".substring(0,part.length) && part!="a/p".substring(0,part.length)) {
            ampstr="";
            }
         else if (part=="am/pm" || part=="a/p") {
            thisformat.operators.push(scfn.commands.date);
            thisformat.operands.push(ampmstr);
            ampmstr = "";
            }
         continue;
         }
      if (ch=="#" || ch=="0" || ch=="?") { // placeholder
         if (integerpart) {
            sectioninfo.integerdigits++;
            if (sectioninfo.commas) { // comma inside of integer placeholders
               sectioninfo.thousandssep = 1; // any number is thousands separator
               sectioninfo.commas = 0; // reset count of "thousand" factors
               }
            lastwasinteger = 1;
            thisformat.operators.push(scfn.commands.integer_placeholder);
            thisformat.operands.push(ch);
            }
         else {
            sectioninfo.fractiondigits++;
            lastwasinteger = 1;
            thisformat.operators.push(scfn.commands.fraction_placeholder);
            thisformat.operands.push(ch);
            }
         }
      else if (ch==".") { // decimal point
         lastwasinteger = 0;
         thisformat.operators.push(scfn.commands.decimal);
         thisformat.operands.push(ch);
         integerpart = 0;
         }
      else if (ch=='$') { // currency char
         lastwasinteger = 0;
         thisformat.operators.push(scfn.commands.currency);
         thisformat.operands.push(ch);
         }
      else if (ch==",") {
         if (lastwasinteger) {
            sectioninfo.commas++;
            }
         else {
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch);
            }
         }
      else if (ch=="%") {
         lastwasinteger = 0;
         sectioninfo.percent++;
         thisformat.operators.push(scfn.commands.copy);
         thisformat.operands.push(ch);
         }
      else if (ch=='"') {
         lastwasinteger = 0;
         inquote = 1;
         quotestr = "";
         }
      else if (ch=='[') {
         lastwasinteger = 0;
         inbracket = 1;
         bracketstr = "";
         }
      else if (ch=='\\') {
         lastwasslash = 1;
         lastwasinteger = 0;
         }
      else if (ch=='*') {
         lastwasasterisk = 1;
         lastwasinteger = 0;
         }
      else if (ch=='_') {
         lastwasunderscore = 1;
         lastwasinteger = 0;
         }
      else if (ch==";") {
         section++; // start next section
         thisformat.sectioninfo[section] = {}; // create a new section
         sectioninfo = thisformat.sectioninfo[section]; // get reference to info for current section
         sectioninfo.sectionstart = 1 + thisformat.operators.length; // remember where it starts
         sectioninfo.integerdigits = 0; // number of integer-part placeholders
         sectioninfo.fractiondigits = 0; // fraction placeholders
         sectioninfo.commas = 0; // commas encountered, to handle scaling
         sectioninfo.percent = 0; // times to scale by 100
         integerpart = 1; // reset for new section
         lastwasinteger = 0;
         thisformat.operators.push(scfn.commands.section);
         thisformat.operands.push(ch);
         }
      else if (ch.toLowerCase()=="g") {
         ingeneral = 1;
         lastwasinteger = 0;
         }
      else if (ch.toLowerCase()=="a") {
         ampmstr = ch;
         lastwasinteger = 0;
         }
      else if ("dmyhHs".indexOf(ch)>=0) {
         indate = ch;
         }
      else {
         lastwasinteger = 0;
         thisformat.operators.push(scfn.commands.copy);
         thisformat.operands.push(ch);
         }
      }

   if (indate) { // last char was part of unsaved date placeholder
      thisformat.operators.push(scfn.commands.date);
      thisformat.operands.push(indate);
      sectioninfo.hasdate = 1;
      }

   return;

   }


/* *******************

 bracketdata = SocialCalc.FormatNumber.parse_format_bracket(bracketstr)

 Takes a bracket contents (e.g., "RED", ">10") and returns an operator and operand

 bracketdata->{}
    .operator
    .operand

************************* */

SocialCalc.FormatNumber.parse_format_bracket = function(bracketstr) {

   var scfn = SocialCalc.FormatNumber;
   var scc = SocialCalc.Constants;

   var bracketdata={};
   var parts;

   if (bracketstr.charAt(0)=='$') { // currency
      bracketdata.operator = scfn.commands.currency;
      parts=bracketstr.match(/^\$(.+?)(\-.+?){0,1}$/);
      if (parts) {
         bracketdata.operand = parts[1] || scc.FormatNumber_defaultCurrency || '$';
         }
      else {
         bracketdata.operand = bracketstr.substring(1) || scc.FormatNumber_defaultCurrency || '$';
         }
      }
   else if (bracketstr=='?$') {
      bracketdata.operator = scfn.commands.currency;
      bracketdata.operand = '[?$]';
      }
   else if (scfn.allowedcolors[bracketstr.toUpperCase()]) {
      bracketdata.operator = scfn.commands.color;
      bracketdata.operand = scfn.allowedcolors[bracketstr.toUpperCase()];
      }
   else if (parts=bracketstr.match(/^style=([^"]*)$/)) { // [style=...]
      bracketdata.operator = scfn.commands.style;
      bracketdata.operand = parts[1];
      }
   else if (bracketstr==",") {
      bracketdata.operator = scfn.commands.separator;
      bracketdata.operand = bracketstr;
      }
   else if (scfn.alloweddates[bracketstr.toUpperCase()]) {
      bracketdata.operator = scfn.commands.date;
      bracketdata.operand = scfn.alloweddates[bracketstr.toUpperCase()];
      }
   else if (parts=bracketstr.match(/^[<>=]/)) { // comparison operator
      parts=bracketstr.match(/^([<>=]+)(.+)$/); // split operator and value
      bracketdata.operator = scfn.commands.comparison;
      bracketdata.operand = parts[1]+":"+parts[2];
      }
   else { // unknown bracket
      bracketdata.operator = scfn.commands.copy;
      bracketdata.operand = "["+bracketstr+"]";
      }

   return bracketdata;

   }

/* *******************

 juliandate = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(year, month, day)

 From: http://aa.usno.navy.mil/faq/docs/JD_Formula.html
 Uses: Fliegel, H. F. and van Flandern, T. C. (1968). Communications of the ACM, Vol. 11, No. 10 (October, 1968).
 Translated from the FORTRAN

      I= YEAR
      J= MONTH
      K= DAY
C
      JD= K-32075+1461*(I+4800+(J-14)/12)/4+367*(J-2-(J-14)/12*12)
     2    /12-3*((I+4900+(J-14)/12)/100)/4

************************* */

SocialCalc.FormatNumber.convert_date_gregorian_to_julian = function(year, month, day) {

   var juliandate;

   juliandate = day-32075+SocialCalc.intFunc(1461*(year+4800+SocialCalc.intFunc((month-14)/12))/4);
   juliandate += SocialCalc.intFunc(367*(month-2-SocialCalc.intFunc((month-14)/12)*12)/12);
   juliandate = juliandate - SocialCalc.intFunc(3*SocialCalc.intFunc((year+4900+SocialCalc.intFunc((month-14)/12))/100)/4);

   return juliandate;

   }


/* *******************

 ymd = SocialCalc.FormatNumber.convert_date_julian_to_gregorian(juliandate)

 ymd->{}
    .year
    .month
    .day

 From: http://aa.usno.navy.mil/faq/docs/JD_Formula.html
 Uses: Fliegel, H. F. and van Flandern, T. C. (1968). Communications of the ACM, Vol. 11, No. 10 (October, 1968).
 Translated from the FORTRAN

************************* */

SocialCalc.FormatNumber.convert_date_julian_to_gregorian = function(juliandate) {

   var L, N, I, J, K;

   L = juliandate+68569;
   N = Math.floor(4*L/146097);
   L = L-Math.floor((146097*N+3)/4);
   I = Math.floor(4000*(L+1)/1461001);
   L = L-Math.floor(1461*I/4)+31;
   J = Math.floor(80*L/2447);
   K = L-Math.floor(2447*J/80);
   L = Math.floor(J/11);
   J = J+2-12*L;
   I = 100*(N-49)+I+L;

   return {year:I, month:J, day:K};

   }

SocialCalc.intFunc = function(n) {
   if (n < 0) {
      return -Math.floor(-n);
      }
   else {
      return Math.floor(n);
      }
   }

if (typeof global != 'undefined') var window = global;
if (typeof SocialCalc != 'undefined' && typeof module != 'undefined') module.exports = SocialCalc;
